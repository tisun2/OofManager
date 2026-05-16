Import-Module (Join-Path (Split-Path $script:MyInvocation.MyCommand.Path) "Microsoft.PowerApps.RestClientModule.psm1") -NoClobber #-Force
Import-Module (Join-Path (Split-Path $script:MyInvocation.MyCommand.Path) "Microsoft.PowerApps.AuthModule.psm1") -NoClobber #-Force



function Add-ConnectorToBusinessDataGroup
{
    <#
    .SYNOPSIS
    Sets connector to the business data group of data loss policy.
    .DESCRIPTION
    The Add-ConnectorToBusinessDataGroup set connector to the business data group of DLP depending on parameters.
    Use Get-Help Add-ConnectorToBusinessDataGroup -Examples for more detail.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER ConnectorName
    The Connector's identifier.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER ApiVersion
    The api version to call with. Default 2018-01-01.
    .EXAMPLE
    Add-ConnectorToBusinessDataGroup -PolicyName e25a94b2-3111-468e-9125-3d3db3938f13 -ConnectorName shared_office365users
    Sets the connector to BusinessData group of policyname e25a94b2-3111-468e-9125-3d3db3938f13.
    .EXAMPLE
    Add-ConnectorToBusinessDataGroup -EnvironmentName Default-02c201b0-db76-4a6a-b3e1-a69202b479e6 -PolicyName e25a94b2-3111-468e-9125-3d3db3938f13 -ConnectorName shared_office365users
    Sets the connector to BusinessData group of policyname e25a94b2-3111-468e-9125-3d3db3938f13 in environment Default-02c201b0-db76-4a6a-b3e1-a69202b479e6.
    #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2018-01-01"
    )
    process
    {
        if([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies/{policyName}?api-version={apiVersion}"
        }
        else
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/apiPolicies/{policyName}?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
        }

        $route = $route | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $policy = InvokeApi -Method Get -Route $route -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        $existingConnector = $policy.properties.definition.apiGroups.hbi.apis | where { ($_.id -split "/apis/")[1] -eq $ConnectorName }
        if($existingConnector -ne $null)
        {
            Write-Error "Connector already exists in business data group"
            return $null
        }
        $environments = Get-AdminPowerAppEnvironment
        $connector = $environments[0] | Get-PowerAppConnector -ConnectorName $ConnectorName `
            | %{ New-Object -TypeName PSObject -Property @{ id = $_.connectorId; name = $_.internal.properties.displayName; type = $_.internal.type } }

        if($connector -eq $null)
        {
            Write-Error "No connector with specified name found"
            return $null
        }

        #Add it to the hbi object of policy
        $policy.properties.definition.apiGroups.hbi.apis += $connector

        #remove from lbi object of policy
        $lbiWithoutProvidedConnector = $policy.properties.definition.apiGroups.lbi.apis | Where-Object { $_.id -ne $connector.id }

        if ($lbiWithoutProvidedConnector -eq $null)
        {
            $lbiWithoutProvidedConnector =  @()
        }

        $policy.properties.definition.apiGroups.lbi.apis = [Array]$lbiWithoutProvidedConnector

        #API Call
        $setConnectorResult = InvokeApiNoParseContent -Method PUT -Body $policy -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        CreateHttpResponse($setConnectorResult)
    }
}

function Remove-ConnectorFromBusinessDataGroup
{
     <#
    .SYNOPSIS
    Removes connector to the business data group of data loss policy.
    .DESCRIPTION
    The Remove-ConnectorFromBusinessDataGroup removes connector from the business data group of DLP depending on parameters.
    Use Get-Help Remove-ConnectorFromBusinessDataGroup -Examples for more detail.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER ConnectorName
    The Connector's identifier.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER ApiVersion
    The api version to call with. Default 2018-01-01.
    .EXAMPLE
    Remove-ConnectorFromBusinessDataGroup -PolicyName e25a94b2-3111-468e-9125-3d3db3938f13 -ConnectorName shared_office365users
    Removes the connector from BusinessData group of policyname e25a94b2-3111-468e-9125-3d3db3938f13.
    .EXAMPLE
    Remove-ConnectorFromBusinessDataGroup -EnvironmentName Default-02c201b0-db76-4a6a-b3e1-a69202b479e6 -PolicyName e25a94b2-3111-468e-9125-3d3db3938f13 -ConnectorName shared_office365users
    Removes the connector from BusinessData group of policyname e25a94b2-3111-468e-9125-3d3db3938f13 in environment Default-02c201b0-db76-4a6a-b3e1-a69202b479e6.
    #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2018-01-01"
    )
    process
    {
        if([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies/{policyName}?api-version={apiVersion}"
        }
        else
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/apiPolicies/{policyName}?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
        }

        $route = $route | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $policy = InvokeApi -Route $route -Method Get -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        $existingConnector = $policy.properties.definition.apiGroups.lbi.apis | where { ($_.id -split "/apis/")[1] -eq $ConnectorName }
        if($existingConnector -ne $null)
        {
            Write-Error "Connector already exists in non-business data group"
            return $null
        }
        $environments = Get-AdminPowerAppEnvironment
        $connector = $environments[0] | Get-PowerAppConnector -ConnectorName $ConnectorName `
            | %{ New-Object -TypeName PSObject -Property @{ id = $_.connectorId; name = $_.internal.properties.displayName; type = $_.internal.type } }

        if($connector -eq $null)
        {
            Write-Error "No connector with specified name found"
            return $null
        }

        #Add it to the lbi object of policy
        $policy.properties.definition.apiGroups.lbi.apis += $connector
        #remove from hbi object of policy
        $hbiWithoutProvidedConnector = $policy.properties.definition.apiGroups.hbi.apis | Where-Object { $_.id -ne $connector.id }

        if ($hbiWithoutProvidedConnector -eq $null)
        {
            $hbiWithoutProvidedConnector = @()
        }

        $policy.properties.definition.apiGroups.hbi.apis = [Array]$hbiWithoutProvidedConnector

        #API Call
        $setConnectorResult = InvokeApiNoParseContent -Method PUT -Body $policy -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        CreateHttpResponse($setConnectorResult)
    }
}

function Get-AppQuarantineState
{
 <#
 .SYNOPSIS
 Gets the app quarantine state.
 .DESCRIPTION
 The Get-AppQuarantineState cmdlet to get app quarantine state.
 .PARAMETER EnvironmentName
 Environment id of the PowerApp.
 .PARAMETER AppName
 App id of the PowerApp to operate on.
 .PARAMETER ApiVersion
 PowerApps api version date, defaults to "2021-02-01".
 .EXAMPLE
 Get-AppQuarantineState -EnvironmentName a5675ce2-99a8-4b25-9272-6abaf37660fe -AppName c3dba9c8-0f42-4c88-8110-04b582f20735
Return:
    Quarantined
 #>
    param
    (
		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

		[Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/getAppQuarantineState?api-version={apiVersion}" `
		| ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-AppAsQuarantined
{
<#
 .SYNOPSIS
 Sets the app quarantine state to quarantined.
 .DESCRIPTION
 The Set-AppAsQuarantined changes the quarantine state of a PowerApp to quarantined.
 .PARAMETER EnvironmentName
 Environment id of the PowerApp.
 .PARAMETER AppName
 App id of the PowerApp to operate on.
 .PARAMETER ApiVersion
 The PowerApps api version date, defaults to "2021-02-01".
 .EXAMPLE
	Set-AppAsQuarantined -EnvironmentName a5675ce2-99a8-4b25-9272-6abaf37660fe -AppName c3dba9c8-0f42-4c88-8110-04b582f20735
 #>
    param
    (
		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

		[Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
            "QuarantineStatus" = "quarantined"
        }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setAppQuarantineState?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName `
		| ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        return InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-AppAsUnquarantined
{
<#
 .SYNOPSIS
 Sets the app quarantine state to unquarantined.
 .DESCRIPTION
 The Set-AppAsUnquarantined changes the quarantine state of a PowerApp to unquarantined.
 .PARAMETER EnvironmentName
 Environment id of the PowerApp.
 .PARAMETER AppName
 App id of the PowerApp to operate on.
 .PARAMETER ApiVersion
 The PowerApps api version date, defaults to "2021-02-01".
 .EXAMPLE
	Set-AppAsUnquarantined -EnvironmentName a5675ce2-99a8-4b25-9272-6abaf37660fe -AppName c3dba9c8-0f42-4c88-8110-04b582f20735
 #>
    param
    (
		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

		[Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
            "QuarantineStatus" = "unquarantined"
        }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setAppQuarantineState?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName `
		| ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        return InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-TenantSettings
{
 <#
 .SYNOPSIS
 Get tenant settings
 .DESCRIPTION
 The Get-TenantSettings cmdlet to get tenant settings.
 Use Get-Help Get-TenantSettings -Examples for more detail.
 .PARAMETER ApiVersion
 The api version to call with. Default 2016-11-01
 .EXAMPLE
 Get-TenantSettings
Return:
    @{walkMeOptOut=False}
 #>
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/listTenantSettings?api-version={apiVersion}"

        return InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-TenantSettings
{
 <#
 .SYNOPSIS
 Set tenant settings
 .DESCRIPTION
 The Set-TenantSettings cmdlet to update tenant settings.
 Use Get-Help Set-TenantSettings -Examples for more detail.
 .PARAMETER RequestBody
 Tenant settings to be updated
 .PARAMETER ApiVersion
 The api version to call with. Default 2016-11-01
 .EXAMPLE
    $requestBody = @{
        WalkMeOptOut = $true
    }

    Set-TenantSettings -RequestBody $requestBody

Return:
    @{walkMeOptOut=True}
 #>
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$RequestBody,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/updateTenantSettings?api-version={apiVersion}"

        return InvokeApi -Method POST -Route $route -Body $RequestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-AllowedConsentPlans
{
 <#
 .SYNOPSIS
 Removes all consent plans of the specified types from the tenant and blocks consent plans of those type from being created within the tenant.
 .DESCRIPTION
 The Remove-AllowedConsentPlans cmdlet to remove consent plans of the specified types from the tenant and block consent plans of those
 type from being created within the tenant. Note that this will only impact the types specified, all others will be unaffected.
 Use Get-Help Remove-AllowedConsentPlans -Examples for more detail.

 .PARAMETER Types
 The types of consent plans that should be removed. Valid options are Internal, Viral.
 .PARAMETER Prompt
 Prompt remove consent plans confirmation
 .PARAMETER ApiVersion
 The api version to call with. Default 2016-11-01
 .EXAMPLE
    Remove-AllowedConsentPlans -Types @("Internal", "Viral")
 #>
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Internal", "Viral", IgnoreCase = $true)]
        [string[]] $Types,

        [Parameter(Mandatory = $false)]
        [bool]$Prompt = $true,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        if ($Prompt)
        {
            $question = Read-Host "Are you sure you want to proceed? This command will remove all consent plans of the specified type(s) from all users within your tenant. Enter 'y' to proceed."

            if ($question -ne 'y')
            {
                return
            }
        }

        foreach ($type in $Types)
        {
            Write-Verbose "Removing all consent plans of type $type."

            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/allowedConsentPlans/" + $type + "?api-version={apiVersion}"

            InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }
    }
}

function Add-AllowedConsentPlans
{
 <#
 .SYNOPSIS
 Allows consent plans of the specified types to be created within the tenant.
 .DESCRIPTION
 The Add-AllowedConsentPlans cmdlet to allow creation of consent plans of the specified types within the tenant.
 Note that this will only impact the types specified, all others will be unaffected.
 Use Get-Help Add-AllowedConsentPlans -Examples for more detail.

  .PARAMETER Types
 The types of consent plans that should become allowed. Valid options are Internal, Viral.
 .PARAMETER ApiVersion
 The api version to call with. Default 2016-11-01
 .EXAMPLE
    Add-AllowedConsentPlans -Types @("Internal", "Viral")
 #>
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Internal", "Viral", IgnoreCase = $true)]
        [string[]] $Types,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $requestBody = @{
            "types" = $Types
        }

        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/allowedConsentPlans?api-version={apiVersion}"

        return InvokeApi -Method POST -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-AllowedConsentPlans
{
 <#
 .SYNOPSIS
 Gets the types of consent plans that are allowed within the tenant.
 .DESCRIPTION
 The Get-AllowedConsentPlans cmdlet to get all of the types of consent plans that are allowed within the tenant.
 Use Get-Help Get-AllowedConsentPlans -Examples for more detail.

 .PARAMETER ApiVersion
 The api version to call with. Default 2016-11-01
 .EXAMPLE
    Get-AllowedConsentPlans
 #>
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/allowedConsentPlans?api-version={apiVersion}"

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Add-ConnectorsToPolicy
{
    <#
    .SYNOPSIS
    Adds the given connectors to the given group of the DLP policy.
    .DESCRIPTION
    The Add-ConnectorsToPolicy takes the connectors as a list of objects and adds them to the specified group of the given DLP policy.
	If any of the connectors are already classified in a different group, they are moved to the specified group.
    Use Get-Help Add-ConnectorsToPolicy -Examples for more detail.
    .PARAMETER PolicyName
    The policy name (GUID).
    .PARAMETER Classification
    The name of the group to which the connectors should be added. Allowed values are 'General', 'Confidential', and 'Blocked' (for Non-business, Business, and Blocked, respectively).
    .PARAMETER Connectors
    The list of connectors to add to the policy. Each entry in the list is required to have 3 properties:
	  - id: The unique identifier of the connector. Examples:
	    - /providers/Microsoft.PowerApps/apis/shared_msnweather (for connectors)
	    - providers/Microsoft.ProcessSimple/operationGroups/DesktopFlow.Scripting (for Desktop Flow modules)
	  - type: The type of the connector:
	    - providers/Microsoft.PowerApps/apis (for connectors)
	    - Microsoft.ProcessSimple/operationGroups (for Desktop Flow modules)
	  - name: The friendly name of the connector, e.g. "MSN Weather" for the MSN Weather connector or "Active Directory" for the Active Directory Desktop Flow module
    .PARAMETER ApiVersion
    The api version to call with. The default is 2018-01-01.
    .EXAMPLE
	$connectorsToClassifyAsBusiness = @([pscustomobject]@{
		id = "/providers/Microsoft.PowerApps/apis/shared_msnweather"
		name = "MSN Weather"
		type = "providers/Microsoft.PowerApps/apis"
	})
    Add-ConnectorsToPolicy -PolicyName 7b914a18-ad8b-4f15-8da5-3155c77aa70a -Connectors $connectorsToClassifyAsBusiness -Classification Confidential
    Adds the MSN Weather connector to the Business group of the policy. If the connector is already in the policy, moves it to the Business group.
    #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string][ValidateSet("General", "Confidential", "Blocked")]$Classification,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [object[]]$Connectors,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2018-01-01"
    )
    process
    {
		$dlpPolicy = Get-DlpPolicy -PolicyName $PolicyName
		if ($dlpPolicy -eq $null)
		{
			Write-Error "No policy was found with the given name."
			return $null
		}

		if ($Connectors -eq $null -or $Connectors.Count -eq 0)
		{
			Write-Error "The connector list cannot be null or empy."
			return $null
		}

		$connectorsToAdd = @()
        foreach ($connector in $Connectors)
        {
            Write-Verbose "Processing connector $connector"

			if ($connector.id -eq $null -or $connector.name -eq $null -or $connector.type -eq $null)
			{
				Write-Error "The connector $connector is missing a required property. See the description of the 'Connectors' parameter for details."
				return $null
			}

			$connectorIdSegments = ($connector.id -split "/")
			$connectorName = $connectorIdSegments[$connectorIdSegments.Length - 1]

			# If the connector is already in the policy, remove it. It will be added later.
			foreach ($existingConnectorGroup in $dlpPolicy.connectorGroups)
			{
				$existingConnectorInPolicy = $existingConnectorGroup.connectors | where { $_.id -like "*$connectorName" }
				$existingConnectorGroupClassification = $existingConnectorGroup.classification
				if ($existingConnectorInPolicy -ne $null)
				{
					if ($existingConnectorGroup.classification -eq $Classification)
					{
						Write-Verbose "Connector $connectorName is already in the specified group."
					}
					else
					{
						Write-Verbose "Connector $connectorName is in group '$existingConnectorGroupClassification'. It will be moved to '$Classification'."
					}
					$existingConnectorGroup.connectors = $existingConnectorGroup.connectors | where { $_.id -notlike "*$connectorName" }
				}
			}

			$connectorsToAdd += [pscustomobject]@{
				id = $connector.id
				name = $connector.name
				type = $connector.type
			}
        }

		$connectorGroupToAddConnectorsTo = $dlpPolicy.connectorGroups | where { $_.classification -eq $Classification }
		$connectorGroupToAddConnectorsTo.connectors += $connectorsToAdd

		Set-DlpPolicy -PolicyName $PolicyName -UpdatedPolicy $dlpPolicy
    }
}

function Add-CustomConnectorToPolicy
{
    <#
    .SYNOPSIS
    Adds a custom connector to the given group.
    .DESCRIPTION
    The Add-CustomConnectorToPolicy adds a custom connector to a specific group of a DLP policy depending on parameters.
    Use Get-Help Add-CustomConnectorToPolicy -Examples for more detail.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER GroupName
    The name of the group to add the connector to, lbi or hbi.
    .PARAMETER ConnectorName
    The Custom Connector's name.
    .PARAMETER ConnectorId
    The Custom Connector's ID.
    .PARAMETER ConnectorType
    The Custom Connector's type.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER ApiVersion
    The api version to call with. Default 2018-01-01.
    .EXAMPLE
    Add-CustomConnectorToPolicy -PolicyName 7b914a18-ad8b-4f15-8da5-3155c77aa70a -ConnectorName BloopBlop -ConnectorId /providers/Microsoft.PowerApps/apis/BloopBlop -ConnectorType Microsoft.PowerApps/apis -GroupName hbi
    Adds the custom connector 'BloopBlop' to BusinessData group of policy name 7b914a18-ad8b-4f15-8da5-3155c77aa70a.
    .EXAMPLE
    Add-CustomConnectorToPolicy -EnvironmentName Default-02c201b0-db76-4a6a-b3e1-a69202b479e6 -PolicyName 7b914a18-ad8b-4f15-8da5-3155c77aa70a -ConnectorName BloopBlop -ConnectorId /providers/Microsoft.PowerApps/apis/BloopBlop -ConnectorType Microsoft.PowerApps/apis -GroupName hbi
    Adds the custom connector 'BloopBlop' to BusinessData group of policy name 7b914a18-ad8b-4f15-8da5-3155c77aa70a in environment Default-02c201b0-db76-4a6a-b3e1-a69202b479e6.
    #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string][ValidateSet("lbi", "hbi")]$GroupName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorId,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorType,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2018-01-01"
    )
    process
    {
        if([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies/{policyName}?api-version={apiVersion}"
        }
        else
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/apiPolicies/{policyName}?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
        }

        $route = $route | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $policy = InvokeApi -Method Get -Route $route -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        $connectorJsonLbi = $policy.properties.definition.apiGroups.lbi.apis | where { $_.id -eq $ConnectorId }
        $connectorJsonHbi = $policy.properties.definition.apiGroups.hbi.apis | where { $_.id -eq $ConnectorId }

        if($connectorJsonLbi -eq $null -and $connectorJsonHbi -eq $null)
        {
            $customConnectorJson = @{
                id = $ConnectorId
                name = $ConnectorName
                type = $ConnectorType
            }

            if ($GroupName -eq "hbi")
            {
                #Add it to the hbi object of policy
                $policy.properties.definition.apiGroups.hbi.apis += $customConnectorJson
            }
            else
            {
                #Add it to the lbi object of policy
                $policy.properties.definition.apiGroups.lbi.apis += $customConnectorJson
            }

            #APi Call
            $setConnectorResult = InvokeApiNoParseContent -Method PUT -Body $policy -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            CreateHttpResponse($setConnectorResult)
        }
        else
        {
            if($connectorJsonLbi -eq $null)
            {
                Write-Error "The given connector is already present in the hbi group."
            }
            else
            {
                Write-Error "The given connector is already present in the lbi group."
            }
            return $null
        }
    }
}

function Remove-CustomConnectorFromPolicy
{
    <#
    .SYNOPSIS
    Deletes a custom connector from the given DLP policy.
    .DESCRIPTION
    The Delete-CustomConnectorFromPolicy deletes a custom connector from the specific DLP policy.
    Use Get-Help Delete-CustomConnectorFromPolicy -Examples for more detail.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER ConnectorName
    The connector's identifier.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER ApiVersion
    The api version to call with. Default 2018-01-01.
    .EXAMPLE
    Delete-CustomConnectorFromPolicy -PolicyName 7b914a18-ad8b-4f15-8da5-3155c77aa70a -ConnectorName shared_office365users
    Deletes the custom connector 'shared_office365users' from the DLP policy of policy name 7b914a18-ad8b-4f15-8da5-3155c77aa70a.
    .EXAMPLE
    Delete-CustomConnectorFromPolicy -EnvironmentName Default-02c201b0-db76-4a6a-b3e1-a69202b479e6 -PolicyName 7b914a18-ad8b-4f15-8da5-3155c77aa70a -ConnectorName shared_office365users
    Deletes the custom connector 'shared_office365users' from the DLP policy of policy name 7b914a18-ad8b-4f15-8da5-3155c77aa70a in environment Default-02c201b0-db76-4a6a-b3e1-a69202b479e6.
    #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2018-01-01"
    )
    process
    {
        if([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies/{policyName}?api-version={apiVersion}"
        }
        else
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/apiPolicies/{policyName}?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
        }

        $route = $route | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $policy = InvokeApi -Method Get -Route $route -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        $connectorJsonLbi = $policy.properties.definition.apiGroups.lbi.apis | where { ($_.id -split "/apis/")[1] -eq $ConnectorName }
        $connectorJsonHbi = $policy.properties.definition.apiGroups.hbi.apis | where { ($_.id -split "/apis/")[1] -eq $ConnectorName }

        if ($connectorJsonLbi -eq $null) {
            $connectorJsonLbi = $policy.properties.definition.apiGroups.lbi.apis | where { $_.id -notcontains "apis/shared_"} | where { $_.id -eq $ConnectorName }
        }

        if ($connectorJsonHbi -eq $null) {
            $connectorJsonHbi = $policy.properties.definition.apiGroups.hbi.apis | where { $_.id -notcontains "apis/shared_"} | where { $_.id -eq $ConnectorName }
        }

        if($connectorJsonLbi -eq $null -and $connectorJsonHbi -eq $null)
        {
            Write-Error "The given connector is not in the policy."
            return $null
        }
        else
        {
            if($connectorJsonLbi -eq $null)
            {
                #remove from hbi object of policy
                $hbiWithoutProvidedConnector = $policy.properties.definition.apiGroups.hbi.apis -ne $connectorJsonHbi
                $policy.properties.definition.apiGroups.hbi.apis =  $hbiWithoutProvidedConnector

                #APi Call
                $removeConnectorResult = InvokeApiNoParseContent -Method PUT -Body $policy -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                CreateHttpResponse($removeConnectorResult)
            }
            else
            {
                #remove from lbi object of policy
                $lbiWithoutProvidedConnector = $policy.properties.definition.apiGroups.lbi.apis -ne $connectorJsonLbi
                $policy.properties.definition.apiGroups.lbi.apis =  $lbiWithoutProvidedConnector

                #APi Call
                $removeConnectorResult = InvokeApiNoParseContent -Method PUT -Body $policy -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                CreateHttpResponse($removeConnectorResult)
            }
        }
    }
}

function Get-AdminPowerAppConnectionReferences
{
 <#
 .SYNOPSIS
 Returns app connection references.
 .DESCRIPTION
 The Get-AdminPowerAppConnectionReferences information about all connections referenced in an input PowerApp.
 Use Get-Help Get-AdminPowerAppConnectionReferences -Examples for more detail.
 .PARAMETER AppName
 PowerApp to list connectors for.
 .PARAMETER EnvironmentName
 Environment where the input PowerApp is located.
 .EXAMPLE
 Get-AdminPowerAppConnectionReferences -EnvironmentName 643268a6-c680-446f-b8bc-a3ebbf98895f -AppName fc947231-728a-4a74-a654-64b0f22a0d71
 Returns all connections referenced in the PowerApp fc947231-728a-4a74-a654-64b0f22a0d71.
 #>
    [CmdletBinding(DefaultParameterSetName="Connector")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Connector", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,
        [Parameter(Mandatory = $true, ParameterSetName = "Connector", ValueFromPipelineByPropertyName = $true)]
        [string]$AppName
    )
    process
    {
        $app = Get-AdminPowerApp -EnvironmentName $EnvironmentName -AppName $AppName

        if ($app.Internal.properties.connectionReferences -ne $null)
        {
            foreach($conRef in $app.Internal.properties.connectionReferences)
            {
                foreach($connection in $conRef)
                {
                    foreach ($connId in ($connection | Get-Member -MemberType NoteProperty).Name)
                    {
                        Get-PowerAppConnector -EnvironmentName $EnvironmentName -ConnectorName ($($connection.$connId).id -split '/apis/')[1]
                    }
                }
            }
        }
        else
        {
            return $app
        }
    }
}

function Get-AdminPowerAppConnectorAction
{
 <#
 .SYNOPSIS
 Returns connector operations.
 .DESCRIPTION
 The Get-AdminPowerAppConnectorAction gets the specified connector operations.
 Use Get-Help Get-AdminPowerAppConnectorAction -Examples for more detail.
 .PARAMETER ConnectorName
 The connector name.
 .PARAMETER EnvironmentName
 The Environment's identifier.
 .PARAMETER ApiVersion
 The api version to call with. Default 2018-01-01.
 .EXAMPLE
 Get-AdminPowerAppConnectorAction -ConnectorName "shared_powerplatformforadmins"
 Returns all operations for the connector.
 #>
    [CmdletBinding(DefaultParameterSetName="Connector")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Connector", ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName = $null,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        if ([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/connectors/{connector}/actions?api-version={apiVersion}"  `
            | ReplaceMacro -Macro "{connector}" -Value $ConnectorName
        }
        else
        {
            $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/environments/{environmentName}/connectors/{connector}/actions?api-version={apiVersion}"  `
            | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
        }

        $response = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        foreach($action in $response.value)
        {
            if ($action.properties.trigger -eq $null)
            {
                CreateConnectorActionObject -ActionObject $action
            }
        }
    }
}

function Get-AdminPowerAppConnector
{
 <#
 .SYNOPSIS
 Returns information about one or more custom connectors.
 .DESCRIPTION
 The Get-AdminConnector looks up information about one or more custom connectors depending on parameters.
 Use Get-Help Get-AdminConnector -Examples for more detail.
 .PARAMETER Filter
 Finds custom connector matching the specified filter (wildcards supported).
 .PARAMETER ConnectorName
 Limit custom connectors returned to those of a specified connector.
 .PARAMETER EnvironmentName
 Limit custom connectors returned to those in a specified environment.
 .PARAMETER CreatedBy
 Limit custom connectors returned to those created by the specified user (email or AAD Principal object id)
 .PARAMETER ApiVersion
 The api version to call with. Default 2017-05-01
 .EXAMPLE
 Get-AdminConnector
 Returns all custom connector from all environments where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
 .EXAMPLE
 Get-AdminConnector *customapi*
 Returns all custom connectors with the text "customapi" in the name/display name from all environments where the calling user is an Environment Admin  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminConnector -CreatedBy foo@bar.onmicrosoft.com
 Returns all apps created by the user with an email of "foo@bar.onmicrosoft.com" from all environment where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminConnector -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f
 Finds custom connectors within the 0fc02431-15fb-4563-a5ab-8211beb2a86f environment
  .EXAMPLE
 Get-AdminConnector -ConnectorName shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698
 Finds all custom connectosr created with name/id shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698 from all environments where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminConnector -ConnectorName shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698 -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f
 Finds connections within the 0fc02431-15fb-4563-a5ab-8211beb2a86f environment that are created against the shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698.
 .EXAMPLE
 Get-AdminConnector *customapi* -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f
 Finds all connections in environment 0fc02431-15fb-4563-a5ab-8211beb2a86f that contain the string "customapi" in their display name.
 #>
    [CmdletBinding(DefaultParameterSetName="Filter")]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "User")]
        [string[]]$Filter,

        [Parameter(Mandatory = $false, ParameterSetName = "Connector", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Connector", ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$CreatedBy,

        [Parameter(Mandatory = $false, ParameterSetName = "Connector")]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$ApiVersion = "2017-05-01"
    )

    process
    {
        # If the connector name is specified, only return connections for that connector
        if (-not [string]::IsNullOrWhiteSpace($ConnectorName))
        {
            $environments = @();

            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $environments += @{
                    EnvironmentName = $EnvironmentName
                }
            }
            else
            {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false
            }

            foreach($environment in $environments)
            {

                $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apis?api-version={apiVersion}" `
                | ReplaceMacro -Macro "{environmentName}" -Value $environment.EnvironmentName;

                $connectionResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                # There is no api endpoint with connector name to get the details for specified connector, so assigning filter to connectorname to just return the specified connector details
                Get-FilteredCustomConnectors -Filter $ConnectorName -CreatedBy $CreatedBy -ConnectorResult $connectionResult
            }
        }
        else
        {
            # If the caller passed in an environment scope, filter the query to only that environment
            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apis?api-version={apiVersion}" `
                | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

                $connectionResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion

                Get-FilteredCustomConnectors -Filter $Filter -CreatedBy $CreatedBy -ConnectorResult $connectionResult
            }
            # otherwise search for the apps acroos all environments for this calling user
            else
            {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false

                foreach($environment in $environments)
                {
                    $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apis?api-version={apiVersion}" `
                    | ReplaceMacro -Macro "{environmentName}" -Value $environment.EnvironmentName;

                    $connectionResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                    Get-FilteredCustomConnectors -Filter $Filter -CreatedBy $CreatedBy -ConnectorResult $connectionResult
                }
            }
        }
    }
}

function Get-DesktopFlowModules
{
	<#
	.SYNOPSIS
	Returns information about the Desktop Flow modules that can be managed by DLP policies.
	.DESCRIPTION
	The Get-DesktopFlowModules cmdlet looks up information about all available Desktop Flow modules that can be added to DLP policies.
	Use Get-Help Get-DesktopFlowModules -Examples for more details.
	.PARAMETER ApiVersion
	The API version to call the service with. The default is "2021-07-01"
	.EXAMPLE
	Get-DesktopFlowModules
	Finds all Desktop Flow modules that can be managed by DLP policies.
	#>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-07-01"
    )
    process
    {
		Write-Verbose "Starting"
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/connectors?`$filter=connectorType eq 'DesktopFlow'" `
        | ReplaceMacro -Macro "{bapEndpoint}" -Value $global:currentSession.bapEndpoint;

        return Invoke-Request -Uri $route -Method GET -Body $null -ParseContent -ThrowOnFailure
    }
}

function Get-AdminPowerAppConnectorRoleAssignment
{
 <#
 .SYNOPSIS
 Returns the connection role assignments for a user or a custom connection. Owner role assignments cannot be deleted without deleting the connection resource.
 .DESCRIPTION
 The Get-AdminConnectorRoleAssignment functions returns all roles assignments for an custom connector or all custom connectors roles assignments for a user (across all of their connections).  A connection's role assignments determine which users have access to the connection for using or building apps and flows and with which permission level (CanUse, CanUseAndShare) .
 Use Get-Help Get-AdminPowerAppConnectorRoleAssignment -Examples for more detail.
 .PARAMETER EnvironmentName
 The connector's environment.
 .PARAMETER ConnectorName
 The connector's identifier.
 .PARAMETER PrincipalObjectId
 The objectId of a user or group, if specified, this function will only return role assignments for that user or group.
 .EXAMPLE
 Get-AdminPowerAppConnectorRoleAssignment
 Returns all role assignments for all custom connectors in all environments
 .EXAMPLE
 Get-AdminPowerAppConnectorRoleAssignment -ConnectorName shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698
 Returns all role assignments for the connector with name shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698 in all environments
 .EXAMPLE
 Get-AdminPowerAppConnectorRoleAssignment -ConnectorName shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698 -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -PrincipalObjectId a1caec2d-8b48-40cc-8eb8-5cf95b445b46
 Returns all role assignments for the user, or group with an principal object id of a1caec2d-8b48-40cc-8eb8-5cf95b445b46 for the custom connector with name shared_customapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698 in environment with name 0fc02431-15fb-4563-a5ab-8211beb2a86f
 #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$PrincipalObjectId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $selectedObjectId = $null

        if (-not [string]::IsNullOrWhiteSpace($ConnectorName))
        {
            if (-not [string]::IsNullOrWhiteSpace($PrincipalObjectId))
            {
                $selectedObjectId = $PrincipalObjectId;
            }
        }

        $pattern = BuildFilterPattern -Filter $selectedObjectId

        #If Both EnvironmentName and ConnectorName is Provided, Get the details of provided connector in provided Environment
        if (-not [string]::IsNullOrWhiteSpace($ConnectorName) -and -not [string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}/permissions?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

            $connectorRoleResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            foreach ($connectorRole in $connectorRoleResult.Value)
            {
                if (-not [string]::IsNullOrWhiteSpace($PrincipalObjectId))
                {
                    if ($pattern.IsMatch($connectorRole.properties.principal.id ) -or
                        $pattern.IsMatch($connectorRole.properties.principal.email) -or
                        $pattern.IsMatch($connectorRole.properties.principal.tenantId))
                    {
                        CreateCustomConnectorRoleAssignmentObject -ConnectorRoleAssignmentObj $connectorRole -EnvironmentName $EnvironmentName
                    }
                }
                else
                {
                    CreateCustomConnectorRoleAssignmentObject -ConnectorRoleAssignmentObj $connectorRole -EnvironmentName $EnvironmentName
                }
            }
        }
        else
        {
            # only if EnvironmentName is provided, get the details of specified environment
            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $connectorsList = Get-AdminPowerAppConnector -EnvironmentName $EnvironmentName -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }
            #only if ConnectorName is provided, get the details of specified ConnectorName in all environments
            elseif(-not [string]::IsNullOrWhiteSpace($ConnectorName))
            {
                $connectorsList = Get-AdminPowerAppConnector -ConnectorName $ConnectorName -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }
            else
            {
                $connectorsList = Get-AdminPowerAppConnector -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }

            foreach($connector in $connectorsList)
            {
                Get-AdminPowerAppConnectorRoleAssignment `
                    -ConnectorName $connector.ConnectorName `
                    -EnvironmentName $connector.EnvironmentName `
                    -PrincipalObjectId $selectedObjectId `
                    -ApiVersion $ApiVersion `
                     -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }
        }
    }
}

function Set-AdminPowerAppConnectorRoleAssignment
{
    <#
    .SYNOPSIS
    Sets permissions to the custom connectors.
    .DESCRIPTION
    The Set-AdminPowerAppConnectorRoleAssignment set up permission to custom connectors depending on parameters.
    Use Get-Help Set-AdminPowerAppConnectorRoleAssignment -Examples for more detail.
    .PARAMETER ConnectorName
    The custom connector's identifier.
    .PARAMETER EnvironmentName
    The connectors's environment.
    .PARAMETER RoleName
    Specifies the permission level given to the connector: CanView, CanViewWithShare, CanEdit. Sharing with the entire tenant is only supported for CanView.
    .PARAMETER PrincipalType
    Specifies the type of principal this connector is being shared with; a user, a security group, the entire tenant.
    .PARAMETER PrincipalObjectId
    If this connector is being shared with a user or security group principal, this field specified the ObjectId for that principal. You can use the Get-UsersOrGroupsFromGraph API to look-up the ObjectId for a user or group in Azure Active Directory.
    .EXAMPLE
   Set-AdminPowerAppConnectorRoleAssignment -ConnectorName shared_testapi.5f0629412a7d1fe83e.5f6f049093c9b7a698 -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -RoleName CanView -PrincipalType User -PrincipalObjectId a9f34b89-b7f2-48ef-a3ca-1c435bc655a0
    Give the specified user CanView permissions to the connector with name shared_testapi.5f0629412a7d1fe83e.5f6f049093c9b7a698
    #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("CanView", "CanViewWithShare", "CanEdit")]
        [string]$RoleName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("User", "Group", "Tenant")]
        [string]$PrincipalType,

        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$PrincipalObjectId = $null,

        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $TenantId = $Global:currentSession.tenantId

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}/modifyPermissions?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        #Construct the body
        $requestbody = $null

        If ($PrincipalType -eq "Tenant")
        {
            $requestbody = @{
                put = @(
                    @{
                        properties = @{
                            roleName = $RoleName
                            principal = @{
                                id = $TenantId
                                tenantId = $TenantId
                            }
                        }
                    }
                )
            }
        }
        else
        {
            $requestbody = @{
                put = @(
                    @{
                        properties = @{
                            roleName = $RoleName
                            principal = @{
                                id = $PrincipalObjectId
                            }
                        }
                    }
                )
            }
        }

        $setConnectorRoleResult = InvokeApi -Method POST -Body $requestbody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($setConnectorRoleResult)
    }
}

function Remove-AdminPowerAppConnectorRoleAssignment
{
 <#
 .SYNOPSIS
 Deletes a connector role assignment record.
 .DESCRIPTION
 The Remove-AdminPowerAppConnectorRoleAssignment deletes the specific connector role assignment
 Use Get-Help Remove-AdminPowerAppConnectorRoleAssignment -Examples for more detail.
 .PARAMETER RoleId
 The id of the role assignment to be deleted.
 .PARAMETER ConnectorName
 The connector name
 .PARAMETER EnvironmentName
 The connector's environment.
 .EXAMPLE
 Remove-AdminPowerAppConnectorRoleAssignment -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -ConnectorName shared_testapi.5f0629412a7d1fe83e.5f6f049093c9b7a698 -RoleId /providers/Microsoft.PowerApps/scopes/admin/environments/0fc02431-15fb-4563-a5ab-8211beb2a86f/apis/shared_testapi.5f0629412a7d1fe83e.5f6f049093c9b7a698/permissions/a9f34b89-b7f2-48ef-a3ca-1c435bc655a0
 Deletes the role assignment with an id of /providers/Microsoft.PowerApps/scopes/admin/environments/0fc02431-15fb-4563-a5ab-8211beb2a86f/apis/shared_testapi.5f0629412a7d1fe83e.5f6f049093c9b7a698/permissions/a9f34b89-b7f2-48ef-a3ca-1c435bc655a0
 #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$RoleId,

        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}/modifyPermissions`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{apiVersion}"  -Value $ApiVersion;

        #Construct the body
        $requestbody = $null

        $requestbody = @{
            delete = @(
                @{
                    id = $RoleId
                }
            )
        }

        $removeResult = InvokeApi -Method POST -Body $requestbody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($removeResult -eq $null)
        {
            return $null
        }

        CreateHttpResponse($removeResult)
    }
}

function Get-AdminGenerateDataverseEnforceReport
{
 <#
 .SYNOPSIS
 Generates report for dataverse security group enforcement
 .DESCRIPTION
 The Get-AdminGenerateDataverseEnforceReport generates a report of users at risk for dataverse security enforcement.
 To retrieve the report, run the same command with -Force false.
 To generate a new report, run the same command with -Force true.
 Use Get-Help Get-AdminGenerateDataverseEnforceReport -Examples for more detail.
 .PARAMETER Force
 Optional boolean to force a new report
 .EXAMPLE
 Get-AdminGenerateDataverseEnforceReport -Force true
 Will start a new report for dataverse security group enforcement. To retrieve the report, run the same command with -Force false.
 #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $false)]
        [bool]$Force = $false,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2022-11-01"
    )

    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/generateDataverseSecurityGroupEnforcementReport?api-version={apiVersion}&force={force}" `
        | ReplaceMacro -Macro "{apiVersion}"  -Value $ApiVersion `
        | ReplaceMacro -Macro "{force}" -Value $Force;

        $reportResult = InvokeApiNoParseContent -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($reportResult.Content -like "Operation:*") {
            $reportResult.Content
        } elseif ($reportResult.Content -like "No users found") {
            Write-Output "No users at risk"
        }else{
            $parsedResponse = ConvertFrom-Json $reportResult.Content
            $parsedResponse
        }
    }
}

function Remove-AdminPowerAppConnector
{
 <#
 .SYNOPSIS
 Deletes the custom connector.
 .DESCRIPTION
 The Remove-AdminPowerAppConnector permanently deletes the custom connector.
 Use Get-Help Remove-AdminPowerAppConnector -Examples for more detail.
 .PARAMETER ConnectorName
 The connector's connector name.
 .PARAMETER EnvironmentName
 The connector's environment.
 .EXAMPLE
 Remove-AdminPowerAppConnector -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -ConnectorName shared_testapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698
 Deletes the custom connector with name shared_testapi2.5f0629412a7d1fe83e.5f6f049093c9b7a698
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$ConnectorName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $removeResult = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($removeResult -eq $null)
        {
            return $null
        }

        CreateHttpResponse($removeResult)
    }
}

function Get-AdminPowerAppsUserDetails
{
 <#
 .SYNOPSIS
 Downloads the user details into specified filepath
 .DESCRIPTION
 The Get-AdminPowerAppsUserDetails downloads the powerApps user details into the specified path file
 Use Get-Help Get-AdminPowerAppsUserDetails -Examples for more detail.
 .PARAMETER UserPrincipalName
 The user principal name
 .PARAMETER OutputFilePath
 The Output FilePath
 .EXAMPLE
 Get-AdminPowerAppsUserDetails -OutputFilePath C:\Users\testuser\userdetails.json
 Donloads the details of calling user into specified path file C:\Users\testuser\userdetails.json
 .EXAMPLE
 Get-AdminPowerAppsUserDetails -UserPrincipalName foo@bar.com -OutputFilePath C:\Users\testuser\userdetails.json
 downloads the details of user with principal name foo@bar.com into specified file path C:\Users\testuser\userdetails.json
 #>
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName = $null,

        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath = $null,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        #Construct the body
        $requestbody = $null

        if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName))
        {
            $requestbody = @{
                userPrincipalName = $UserPrincipalName
            }
        }
        #first post call would just return the status 'ACCEPTED' and Location in Headers
        #keep calling Get Location Uri until the job gets finished and once the job gets finished it returns http 'OK' and SAS uri otherwise 'ACCEPTED'
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/exportHiddenUserData?api-version={apiVersion}";

        #Kick-off the job
        $getUserDataJobKickOffResponse = InvokeApiNoParseContent -Route $route -Method POST -Body $requestbody -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        $statusUrl = $getUserDataJobKickOffResponse.Headers['Location']

        #Wait until job is completed
        if (-not [string]::IsNullOrWhiteSpace($statusUrl))
        {
            while($getJobCompletedResponse.StatusCode -ne 200)
            {
                Start-Sleep -s 3
                $getJobCompletedResponse = InvokeApiNoParseContent -Route $statusUrl -Method GET -ThrowOnFailure
            }

            CreateHttpResponse($getJobCompletedResponse)

            $responseBody = ConvertFrom-Json $getJobCompletedResponse.Content

            try {
                $downloadCsvResponse = Invoke-WebRequest -Uri $responseBody -OutFile $OutputFilePath
                Write-Verbose "Downloaded to specified file"
            } catch {
                Write-Host "Error while downloading"
                $response = $_.Exception.Response
                if ($_.ErrorDetails)
                {
                    $errorResponse = ConvertFrom-Json $_.ErrorDetails;
                    $code = $response.StatusCode
                    $message = $errorResponse.error.message
                    Write-Verbose "Status Code: '$code'. Message: '$message'"
                }
            }
        }
    }
}

#internal, helper function
function Get-FilteredCustomConnectors
{
     param
    (
        [Parameter(Mandatory = $false)]
        [object]$Filter,

        [Parameter(Mandatory = $false)]
        [object]$CreatedBy,

        [Parameter(Mandatory = $false)]
        [object]$ConnectorResult
    )

    $patternCreatedBy = BuildFilterPattern -Filter $CreatedBy
    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($connector in $ConnectorResult.Value)
    {
        if ($patternCreatedBy.IsMatch($connector.properties.createdBy.displayName) -or
            $patternCreatedBy.IsMatch($connector.properties.createdBy.email) -or
            $patternCreatedBy.IsMatch($connector.properties.createdBy.id) -or
            $patternCreatedBy.IsMatch($connector.properties.createdBy.userPrincipalName))
        {
            if ($patternFilter.IsMatch($connector.name) -or
                $patternFilter.IsMatch($connector.properties.displayName))
            {
                CreateCustomConnectorObject -CustomConnectorObj $connector
            }
        }
    }
}

#internal, helper function
function CreateCustomConnectorObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$CustomConnectorObj
    )
    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value $CustomConnectorObj.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorId -Value $CustomConnectorObj.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name ApiDefinitions -Value $CustomConnectorObj.properties.apiDefinitions `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $CustomConnectorObj.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $CustomConnectorObj.properties.createdTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedBy -Value $CustomConnectorObj.properties.createdBy `
        | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedTime -Value $CustomConnectorObj.properties.changedTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $CustomConnectorObj.properties.environment.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $CustomConnectorObj.properties;
}

#internal, helper function
function CreateCustomConnectorRoleAssignmentObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ConnectorRoleAssignmentObj,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName
    )

    If($ConnectorRoleAssignmentObj.properties.principal.type -eq "Tenant")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $ConnectorRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $ConnectorRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $ConnectorRoleAssignmentObj.properties.principal.tenantId `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $ConnectorRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $ConnectorRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value ((($ConnectorRoleAssignmentObj.id -split "/apis/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $EnvironmentName `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $ConnectorRoleAssignmentObj;
    }
    elseif($ConnectorRoleAssignmentObj.properties.principal.type -eq "User")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $ConnectorRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $ConnectorRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $ConnectorRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $ConnectorRoleAssignmentObj.properties.principal.email `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $ConnectorRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $ConnectorRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $ConnectorRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value ((($ConnectorRoleAssignmentObj.id -split "/apis/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $EnvironmentName `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $ConnectorRoleAssignmentObj;
    }
    elseif($ConnectorRoleAssignmentObj.properties.principal.type -eq "Group")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $ConnectorRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $ConnectorRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $ConnectorRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $ConnectorRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $ConnectorRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $ConnectorRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value ((($ConnectorRoleAssignmentObj.id -split "/apis/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $EnvironmentName `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $ConnectorRoleAssignmentObj;
    }
    else {
        return $null
    }
}

function New-AdminPowerAppEnvironment
{
<#
 .SYNOPSIS
 Creates an Environment.
 .DESCRIPTION
 The New-AdminPowerAppEnvironment cmdlet creates a new Environment by the logged in user.
 Api version 2019-05-01 adds the capability to create an environment with a database. The new commandline argument 'ProvisionDatabase'
 acts as switch to indicate we need to create an environment with the database. If the switch is set, LanguageName and CurrencyName arguments
 are mandatory to pass while Templates,SecurityGroupId and DomainName are optional.
 Use Get-Help New-AdminPowerAppEnvironment -Examples for more detail.
 .PARAMETER DisplayName
 The display name of the new Environment.
 .PARAMETER LocationName
 The location of the new Environment. Use Get-AdminPowerAppEnvironmentLocations to see the valid locations.
 .PARAMETER RegionName
 OPTIONAL The region of the new Environment. Use Get-AdminPowerAppEnvironmentLocations to see the valid regions.
 .PARAMETER FirstRelease
 OPTIONAL The switch to select a FirstRelease island within the selected Power Platform location.
 .PARAMETER EnvironmentSku
 The Environment type (Trial, Sandbox, Production, SubscriptionBasedTrial, Teams or Developer).
 .PARAMETER ProvisionDatabase
 The switch to provision Cds database along with creating the environment. If set, LanguageName and CurrencyName are mandatory to pass as arguments.
 .PARAMETER CurrencyName
 The default currency for the database, use Get-AdminPowerAppCdsDatabaseCurrencies to get the supported values
 .PARAMETER LanguageName
 The default languages for the database, use Get-AdminPowerAppCdsDatabaseLanguages to get the support values
 .PARAMETER Templates
 The list of templates used for provisioning. If it is null, an empty Common Data Service database will be created
 .PARAMETER TemplateMetadata
 A json object payload customized for the selected templates
 .PARAMETER SecurityGroupId
 The Azure Active Directory security group object identifier to restrict database membership.
 .PARAMETER DomainName
 The domain name.
 .PARAMETER Description
 The environment description (Purpose in instance)
 .PARAMETER WaitUntilFinished
 If set to true, the function will not return until provisioning the database is complete (as either a success or failure)
 .PARAMETER TimeoutInMinutes
 The timeout setting in minutes.
 .EXAMPLE
 New-AdminPowerAppEnvironment -DisplayName 'HQ Apps' -Location unitedstates -EnvironmentSku Trial
 Creates a new Trial Environment in the United States with the display name 'HQ Apps'
 .EXAMPLE
 New-AdminPowerAppEnvironment -DisplayName 'Asia Dev' -Location asia -EnvironmentSku Production
 Creates a new Production Environment in Asia with the display name 'Asia Dev'
 #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$DisplayName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$LocationName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$RegionName,

        [Parameter(Mandatory = $false)]
        [Switch]$FirstRelease,

        [ValidateSet("Trial", "Sandbox", "Production", "SubscriptionBasedTrial", "Teams", "Developer")]
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentSku,

        [Parameter(Mandatory = $false)]
        [Switch]$ProvisionDatabase,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$CurrencyName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$LanguageName,

        [Parameter(Mandatory = $false)]
        [string[]]$Templates,

        [Parameter(Mandatory = $false)]
        [object]$TemplateMetadata,

        [Parameter(Mandatory = $false)]
        [string]$SecurityGroupId = $null,

        [Parameter(Mandatory = $false)]
        [string]$DomainName = $null,

        [Parameter(Mandatory = $false)]
        [string]$Description = $null,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2020-08-01"
    )
    process
    {
        $postEnvironmentUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments`?api-version={apiVersion}&id=/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments";

        if ($EnvironmentSku -eq "Developer" -and !$ProvisionDatabase)
        {
            Write-Error "Developer environments must always include Dataverse provisioning parameters."
            throw
        }
        elseif ($EnvironmentSku -eq "Teams")
        {
            $environment = @{
                location = $LocationName
                properties = @{
                    environmentSku = $EnvironmentSku
                }
            }
        }
        else
        {
            $environment = @{
                location = $LocationName
                properties = @{
                    displayName = $DisplayName
                    environmentSku = $EnvironmentSku
                    description = $Description
                }
            }
        }

        if (-not [string]::IsNullOrEmpty($RegionName))
        {
            $environment.properties["azureRegion"] = $RegionName
        }

        if ($FirstRelease)
        {
            $environment.properties["cluster"] = @{
                category = "FirstRelease"
            }
        }

        if ($ProvisionDatabase)
        {
            if ($CurrencyName -ne $null -and
                $LanguageName -ne $null)
            {
                $environment.properties["linkedEnvironmentMetadata"] = @{
                    baseLanguage = $LanguageName
                    currency = @{
                        code = $CurrencyName
                    }
                    templates = $Templates
                }

                if ($null -ne $TemplateMetadata)
                {
                    $environment.properties["linkedEnvironmentMetadata"] += @{
                       templateMetadata = $TemplateMetadata
                    }
                }

                if (-not [string]::IsNullOrEmpty($SecurityGroupId))
                {
                    $environment.properties["linkedEnvironmentMetadata"] += @{
                       securityGroupId = $SecurityGroupId
                    }

                    if ($EnvironmentSku -eq "Teams")
                    {
                        $environment.properties["connectedGroups"] = @(
                            [pscustomobject]@{
                                id=$SecurityGroupId
                            })
                    }
                }

                if (-not [string]::IsNullOrEmpty($DomainName))
                {
                    $environment.properties["linkedEnvironmentMetadata"] += @{
                        domainName = $DomainName
                    }
                }

                $environment.properties["databaseType"] = "CommonDataService"
            }
            else
            {
                Write-Error "CurrencyName and Language must be passed as arguments."
                throw
            }

            # By default we poll until the CDS database is finished provisioning
            If($WaitUntilFinished)
            {
                $response = InvokeApiNoParseContent -Method POST -Route $postEnvironmentUri -Body $environment -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
                $statusUrl = $response.Headers['Location']

                if ($response.StatusCode -eq 202)
                {
                    $environmentName = ((($statusUrl -split "/environments/")[1]) -split "/")[0]
                    $currentTime = Get-Date -format HH:mm:ss
                    $nextTime = Get-Date -format HH:mm:ss
                    $TimeDiff = New-TimeSpan $currentTime $nextTime

                    #Wait until the environment has been created or the service timeout
                    while((-not [string]::IsNullOrEmpty($statusUrl)) -and ($response.StatusCode -eq 202) -and ($TimeDiff.TotalMinutes -lt $TimeoutInMinutes))
                    {
                        Start-Sleep -s 5
                        $response = InvokeApiNoParseContent -Route $statusUrl -Method GET -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
                        $nextTime = Get-Date -format HH:mm:ss
                        $TimeDiff = New-TimeSpan $currentTime $nextTime
                    }

                    if ($TimeDiff.TotalMinutes -ge $TimeoutInMinutes)
                    {
                        $error = "Provision timeout ($TimeoutInMinutes minutes)."

                        New-Object -TypeName PSObject `
                            | Add-Member -PassThru -MemberType NoteProperty -Name Code -Value 408 `
                            | Add-Member -PassThru -MemberType NoteProperty -Name Description -Value "Request Timeout" `
                            | Add-Member -PassThru -MemberType NoteProperty -Name Headers -Value $response.Headers `
                            | Add-Member -PassThru -MemberType NoteProperty -Name Error -Value $error `
                            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $response;
                    }
                    elseif ($response.StatusCode -eq 200)
                    {
                        # get the environment object for return
                        Get-AdminPowerAppEnvironment -EnvironmentName $environmentName -ReturnCdsDatabaseType $true
                    }
                    else
                    {
                        CreateHttpResponse($response)
                    }
                }
                else
                {
                    CreateHttpResponse($response)
                }
            }
            # optionally the caller can choose to NOT wait until provisioning is complete and get the provisioning status by calling on Get-AdminPowerAppOperationStatus
            else
            {
                $response = InvokeApi -Method POST -Route $postEnvironmentUri -Body $environment -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                CreateHttpResponse($response)
            }
        }
        else
        {
            $response = InvokeApi -Method POST -Route $postEnvironmentUri -ApiVersion $ApiVersion -Body $environment  -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            if ($response.StatusCode -ge 400)
            {
                #Write-Error "An error occured."
                CreateHttpResponse($response)
            }
            else
            {
                CreateEnvironmentObject -EnvObject $response -ReturnCdsDatabaseType $false
            }
        }
    }
}

function Get-AdminPowerAppOperationStatus
{
<#
 .SYNOPSIS
 Get Admin PowerApps operation status by operation Url.
 .DESCRIPTION
 The Get-AdminPowerAppOperationStatus cmdlet gets Admin PowerApps operation status by operation Url.
 If Content.state.id.ToString() -ne "Succeeded", loop and call Get-AdminPowerAppOperationStatus again.
 Use Get-Help Get-AdminPowerAppOperationStatus -Examples for more detail.
 .PARAMETER OperationStatusUrl
 The Url which is returned by API in Headers['Location'] or Headers['Operation-Location'].
 .EXAMPLE
 Get-AdminPowerAppOperationStatus -OperationUrl https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/lifecycleOperations/d14e602b-4bdd-4faa-ad9c-c250a7cccf7b?api-version=2019-05-01
 Get operation status.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$OperationStatusUrl,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $response = InvokeApiNoParseContent -Route $OperationStatusUrl -Method GET -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($response.StatusCode -eq 200)
        {
            # special handling for create environment
            if ($OperationStatusUrl.Contains('/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments'))
            {
                $environmentName = ((($OperationStatusUrl -split "/environments/")[1]) -split "/")[0]

                # get the environment object for return
                $environment = Get-AdminPowerAppEnvironment -EnvironmentName $environmentName -ReturnCdsDatabaseType $true
                if (-not [string]::IsNullOrWhiteSpace($environment.EnvironmentName) -and $environment.CommonDataServiceDatabaseProvisioningState -eq "Succeeded")
                {
                    return $environment
                }
            }
        }

        CreateHttpResponse($response)
    }
}

function Set-AdminPowerAppEnvironmentDisplayName
{
<#
 .SYNOPSIS
 Updates the Environment display name.
 .DESCRIPTION
 The Set-EnvironmentDisplayName cmdlet updates the display name field of the specified Environment.
 Use Get-Help Set-EnvironmentDisplayName -Examples for more detail.
 .PARAMETER EnvironmentName
 The environment name
 .PARAMETER NewDisplayName
 The new display name of the Environment.
 .EXAMPLE
 Set-EnvironmentDisplayName -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb4 -NewDisplayName Applications
 Updates the display name of Environment '8d996ece-8558-4c4e-b459-a51b3beafdb4' to be called 'Applications'.
 .EXAMPLE
 Set-EnvironmentDisplayName -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb4 -NewDisplayName 'Main Organization Apps'
 Updates the display name to be 'Main Organization Apps'
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$NewDisplayName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        $requestBody = @{
            properties = @{
                displayName = $NewDisplayName
            }
        }

        $response = InvokeApi -Method PATCH -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Set-PowerPlatformGenerativeAiSettings
{
<#
.SYNOPSIS
Sets the generative AI settings in an environment.
.DESCRIPTION
The Set-PowerPlatformGenerativeAiSettings cmdlet updates the generative AI setting for a specified environment.
.PARAMETER EnvironmentName
Specifies the ID of the environment you want to modify.
.PARAMETER CrossGeoDataMovement
Enables or disables cross-geo data movement.
.PARAMETER BingSearch
Enables or disables Bing Search capabilities.
.PARAMETER Microsoft365Services
Enables or disables Microsoft 365 services.
.EXAMPLE
Set-PowerPlatformGenerativeAiSettings -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb46 -CrossGeoDataMovement $true -BingSearch $true -Microsoft365Services $true
#>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [bool]$CrossGeoDataMovement,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [bool]$BingSearch,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [bool]$Microsoft365Services,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-04-01"
    )

    process
    {
        # Check if setting is provided
       if ((-not $PSBoundParameters.ContainsKey('CrossGeoDataMovement')) -and (-not $PSBoundParameters.ContainsKey('BingSearch')) -and (-not $PSBoundParameters.ContainsKey('Microsoft365Services'))) {
           Write-Warning "No settings provided. Skipping call."
           return
       }

       $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}`?api-version={apiVersion}" `
       | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

       $requestBody = @{
            properties = @{}
        }
        
       if ($PSBoundParameters.ContainsKey('CrossGeoDataMovement')) {
           $requestBody.properties.Add("copilotPolicies", @{"crossGeoCopilotDataMovementEnabled" = $CrossGeoDataMovement})
       }

       if ($PSBoundParameters.ContainsKey('BingSearch')) {
           $requestBody.properties.Add("bingChatEnabled", $BingSearch)
       }

       if ($PSBoundParameters.ContainsKey('Microsoft365Services')) {
           $requestBody.properties.Add("m365Enabled", $Microsoft365Services)
       }

       $response = InvokeApi -Method PATCH -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

       CreateHttpResponse($response)
    }
}

function Set-AdminPowerAppEnvironmentMakerAnalyticsSettings
{
<#
.SYNOPSIS
Sets the setting to enable/disable analytics for makers in an environment.
.DESCRIPTION
The Set-AdminPowerAppEnvironmentMakerAnalyticsSettings cmdlet updates the enable maker analytics setting for a specified environment.
.PARAMETER EnvironmentName
Specifies the ID of the environment you want to modify.
.PARAMETER EnableAnalyticsForMakers
Enables or disables analytics for makers.
.PARAMETER EnableEUPIForMakerAnalytics
Enables or disables eupi data for analytics for makers.
.EXAMPLE
Set-AdminPowerAppEnvironmentMakerAnalyticsSettings -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb46 -EnableAnalyticsForMakers $true -EnableEUPIForMakerAnalytics $true
#>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [bool]$EnableAnalyticsForMakers,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [bool]$EnableEUPIForMakerAnalytics,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-04-01"
    )

    process
    {
        # Check if setting is provided
       if ((-not $PSBoundParameters.ContainsKey('EnableAnalyticsForMakers')) -and (-not $PSBoundParameters.ContainsKey('EnableEUPIForMakerAnalytics'))) {
           Write-Warning "No settings provided. Skipping call. Skipping call."
           return
       }

       $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}`?api-version={apiVersion}" `
       | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;
        
       $requestBody = @{
        properties = @{}
       }

       if ($PSBoundParameters.ContainsKey('EnableAnalyticsForMakers') -and $EnableAnalyticsForMakers -ne $null) {
           $requestBody.properties.Add("enableAnalyticsForMakers", $EnableAnalyticsForMakers)
       }

       if ($PSBoundParameters.ContainsKey('EnableEUPIForMakerAnalytics') -and $EnableEUPIForMakerAnalytics -ne $null) {
        $requestBody.properties.Add("enableEUPIForMakerAnalytics", $EnableEUPIForMakerAnalytics)
    }

       $response = InvokeApi -Method PATCH -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

       CreateHttpResponse($response)
    }
}

function Set-AdminPowerAppEnvironmentBackupRetentionPeriod
{
<#
 .SYNOPSIS
 Updates the Environment backup retention period.
 .DESCRIPTION
 The Set-AdminPowerAppEnvironmentBackupRetentionPeriod cmdlet updates the backup retention period of the specified Environment.
 Use Get-Help Set-AdminPowerAppEnvironmentBackupRetentionPeriod -Examples for more detail.
 .PARAMETER EnvironmentName
 The environment name
 .PARAMETER NewBackupRetentionPeriodInDays
 The new backup retention period of the Environment in days. Valid values are 7, 14, 21, or 28.
.PARAMETER WaitUntilFinished
 Default is true.  If set to true, then the function will not return until operation completed.
.PARAMETER TimeoutInMinutes
The client timeout setting in minutes.
 .EXAMPLE
 Set-AdminPowerAppEnvironmentBackupRetentionPeriod -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb4 -NewBackupRetentionPeriodInDays 7
 Updates the backup retention period of Environment '8d996ece-8558-4c4e-b459-a51b3beafdb4' to 7 days.
 .EXAMPLE
 Set-AdminPowerAppEnvironmentBackupRetentionPeriod -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb4 -NewBackupRetentionPeriodInDays 28
 Updates the backup retention period of Environment '8d996ece-8558-4c4e-b459-a51b3beafdb4' to 28 days.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true)]
        [ValidateSet(7, 14, 21, 28)]
        [int]$NewBackupRetentionPeriodInDays,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        $retentionTable = @{ 7 = "P7D"; 14 = "P14D"; 21 = "P21D"; 28 = "P28D"}

        $requestBody = @{
            properties = @{
                retentionDetails = @{
                    retentionPeriod = $retentionTable[$NewBackupRetentionPeriodInDays]
                }
            }
        }

        $response = InvokeApi -Method PATCH -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($WaitUntilFinished)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function Set-AdminPowerAppEnvironmentRuntimeState
{
<#
 .SYNOPSIS
 Updates the Environment admin mode.
 .DESCRIPTION
 The Set-AdminPowerAppEnvironmentRuntimeState cmdlet updates the admin mode field of the specified Environment.
 Use Get-Help Set-AdminPowerAppEnvironmentRuntimeState -Examples for more detail.
 .PARAMETER EnvironmentName
 Updates a specific environment.
 .PARAMETER RuntimeState
 The environment runtime state. It can be Enabled, AdminMode, Disabled, Warning. The default is Enabled.
.PARAMETER BackgroundOperationsState
 The environment background operations state. It can be Enabled or Disabled. The default is Enabled.
.PARAMETER WaitUntilFinished
 Default is true.  If set to true, then the function will not return until operation completed.
.PARAMETER TimeoutInMinutes
The client timeout setting in minutes.
 .EXAMPLE
 Set-AdminPowerAppEnvironmentRuntimeState -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb4 -RuntimeState AdminMode
 Set Environment '8d996ece-8558-4c4e-b459-a51b3beafdb4' runtime state to AdminMode.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$RuntimeState,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$BackgroundOperationsState,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        $requestBody = @{
            properties = @{
                LinkedEnvironmentMetadata = @{}
                states = @{
                    runtime = @{
                        id = $RuntimeState
                    }
                }
            }
        }

        if (-not [string]::IsNullOrEmpty($BackgroundOperationsState))
        {
            $requestBody.properties["linkedEnvironmentMetadata"] += @{
                backgroundOperationsState = $BackgroundOperationsState
            }
        }

        $response = InvokeApiNoParseContent -Method PATCH -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($WaitUntilFinished)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function Set-AdminPowerAppEnvironmentGovernanceConfiguration
{
<#
 .SYNOPSIS
 Updates the governance configuration on an existing environment.
 .DESCRIPTION
 The Set-AdminPowerAppEnvironmentGovernanceConfiguration cmdlet updates the governance configuration on the specified environment.
 Use Get-Help Set-AdminPowerAppEnvironmentGovernanceConfiguration -Examples for more details.
 .PARAMETER EnvironmentName
 The name (GUID) of the environment to update.
 .PARAMETER UpdatedGovernanceConfiguration
 The updated governance configuration of the environment.
 .EXAMPLE
  $UpdatedGovernanceConfiguration = [pscustomobject]@{
	protectionLevel = "Basic"
 }
 Set-AdminPowerAppEnvironmentGovernanceConfiguration -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb4 -UpdatedGovernanceConfiguration $UpdatedGovernanceConfiguration
 Sets the protection level of the environment to 'Basic' in the governance configuration.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$UpdatedGovernanceConfiguration,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environmentName}/governanceConfiguration`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        $response = InvokeApi -Method POST -Route $route -Body $UpdatedGovernanceConfiguration -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-AdminPowerAppEnvironmentLocations
{
    <#
    .SYNOPSIS
    Returns all supported environment locations.
    .DESCRIPTION
    The Get-AdminPowerAppEnvironmentLocations cmdlet returns all supported locations to create an environment in PowerApps.
    Use Get-Help Get-AdminPowerAppEnvironmentLocations -Examples for more detail.
    .PARAMETER Filter
    Finds locations matching the specified filter (wildcards supported).
    .EXAMPLE
    Get-AdminPowerAppEnvironmentLocations
    Returns all locations.
    .EXAMPLE
    Get-AdminPowerAppEnvironmentLocations *unitedstates*
    Returns the US location
    #>
    param
    (
        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]$Filter,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )

    $getLocationsUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations?api-version={apiVersion}"

    $locationsResult = InvokeApi -Method GET -ThrowOnFailure -Route $getLocationsUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($location in $locationsResult.Value)
    {
        if ($patternFilter.IsMatch($location.name) -or
            $patternFilter.IsMatch($location.properties.displayName))
        {
            CreateEnvironmentLocationObject -EnvironmentLocationObject $location
        }
    }
}


function Get-AdminPowerAppCdsDatabaseCurrencies
{
    <#
    .SYNOPSIS
    Returns all supported CDS database currencies.
    .DESCRIPTION
    The Get-AdminPowerAppCdsDatabaseCurrencies cmdlet returns all supported database currencies, which is required to provision a new instance.
    Use Get-Help Get-AdminPowerAppCdsDatabaseCurrencies -Examples for more detail.
    .PARAMETER Filter
    Finds currencies matching the specified filter (wildcards supported).
    .PARAMETER LocationName
    The location of the current environment. Use Get-AdminPowerAppEnvironmentLocations to see the valid locations.
    .EXAMPLE
    Get-AdminPowerAppCdsDatabaseCurrencies
    Returns all currencies.
    .EXAMPLE
    Get-AdminPowerAppCdsDatabaseCurrencies *USD*
    Returns the US dollar currency
    #>
    param
    (
        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]$Filter,

        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$LocationName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )

    $getCurrenciesUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/environmentCurrencies?api-version={apiVersion}" `
    | ReplaceMacro -Macro "{location}" -Value $LocationName;

    $currenciesResult = InvokeApi -Method GET -Route $getCurrenciesUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($currency in $currenciesResult.Value)
    {
        if ($patternFilter.IsMatch($currency.name) -or
            $patternFilter.IsMatch($currency.properties.code))
        {
            CreateCurrencyObject -CurrencyObject $currency
        }
    }
}

function Get-AdminPowerAppCdsDatabaseLanguages
{
    <#
    .SYNOPSIS
    Returns all supported CDS database languages.
    .DESCRIPTION
    The Get-AdminPowerAppCdsDatabaseLanguages cmdlet returns all supported database languages, which is required to provision a new instance.
    Use Get-Help Get-AdminPowerAppCdsDatabaseLanguages -Examples for more detail.
    .PARAMETER Filter
    Finds langauges matching the specified filter (wildcards supported).
    .PARAMETER LocationName
    The location of the current environment. Use Get-AdminPowerAppEnvironmentLocations to see the valid locations.
    .EXAMPLE
    Get-AdminPowerAppCdsDatabaseLanguages
    Returns all languages.
    .EXAMPLE
    Get-AdminPowerAppCdsDatabaseLanguages *English*
    Returns all English language options
    #>
    param
    (
        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]$Filter,

        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$LocationName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )

    $getLanguagesUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/environmentLanguages?api-version={apiVersion}" `
    | ReplaceMacro -Macro "{location}" -Value $LocationName;

    $languagesResult = InvokeApi -Method GET -Route $getLanguagesUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($languages in $languagesResult.Value)
    {
        if ($patternFilter.IsMatch($languages.name) -or
            $patternFilter.IsMatch($languages.properties.displayName) -or
            $patternFilter.IsMatch($languages.properties.localizedName))
        {
            CreateLanguageObject -LanguageObject $languages
        }
    }
}

function Get-AdminPowerAppCdsDatabaseTemplates
{
    <#
    .SYNOPSIS
    Returns all supported CDS database templates.
    .DESCRIPTION
    The Get-AdminPowerAppCdsDatabaseTemplates cmdlet returns all supported database templates, which is required to provision a new instance.
    Use Get-Help Get-AdminPowerAppCdsDatabaseTemplates -Examples for more detail.
    .PARAMETER Filter
    Finds templates matching the specified filter (wildcards supported).
    .PARAMETER LocationName
    The location of the current environment. Use Get-AdminPowerAppEnvironmentLocations to see the valid locations.
    .EXAMPLE
    Get-AdminPowerAppCdsDatabaseTemplates -LocationName unitedstates
    Returns all templates for United States.
    .EXAMPLE
    Get-AdminPowerAppCdsDatabaseTemplates *CDS*
    Returns the CDS template
    #>
    param
    (
        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]$Filter,

        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$LocationName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )

    $getTemplatesUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/templates?api-version={apiVersion}" `
    | ReplaceMacro -Macro "{location}" -Value $LocationName;

    $templatesResult = InvokeApi -Method GET -Route $getTemplatesUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($template in $templatesResult.Value)
    {
        if ($patternFilter.IsMatch($template.name) -or
            $patternFilter.IsMatch($template.properties.code))
        {
            CreateTemplateObject -TemplateObject $template
        }
    }
}

function New-AdminPowerAppCdsDatabase
{
    <#
    .SYNOPSIS
    Creates a Common Data Service For Apps database for the specified environment.
    .DESCRIPTION
    The New-AdminPowerAppCdsDatabase cmdlet creates a Common Data Service For Apps database for the specified environment with teh specified default language and currency.
    Use Get-Help New-AdminPowerAppCdsDatabase -Examples for more detail.
    .PARAMETER EnvironmentName
    The environment name
    .PARAMETER CurrencyName
    The default currency for the database, use Get-AdminPowerAppCdsDatabaseCurrencies to get the supported values
    .PARAMETER LanguageName
    The default languages for the database, use Get-AdminPowerAppCdsDatabaseLanguages to get the support values
    .PARAMETER WaitUntilFinished
    Default is true.  If set to true, then the function will not return a value until provisioning the database is complete (as either a success or failure)
    .PARAMETER Templates
    The list of templates that used for provisision. If not provided then an empty Common Data Service database will be created.
    .PARAMETER SecurityGroupId
    The Azure Active Directory security group object identifier to which to restrict database membership.
    .PARAMETER DomainName
    The domain name
    .EXAMPLE
    New-AdminPowerAppCdsDatabase -EnvironmentName 8d996ece-8558-4c4e-b459-a51b3beafdb4 -CurrencyName USD -LanguageName 1033
    Creates a database with the US dollar currency and Environment (US) language
    #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$CurrencyName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$LanguageName,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [string[]]$Templates,

        [Parameter(Mandatory = $false)]
        [string]$SecurityGroupId = $null,

        [Parameter(Mandatory = $false)]
        [string]$DomainName = $null,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$ApiVersion = "2018-01-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environmentName}/provisionInstance`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        $requestBody = @{
            baseLanguage = $LanguageName
            currency = @{
                code = $CurrencyName
            }
            templates = $Templates
        }

        if (-not [string]::IsNullOrEmpty($SecurityGroupId))
        {
            $requestBody += @{
                securityGroupId = $SecurityGroupId
            }
        }
        if (-not [string]::IsNullOrEmpty($DomainName))
        {
            $requestBody += @{
                domainName = $DomainName
            }
        }

        # By default we poll until the CDS database is finished provisioning
        If($WaitUntilFinished)
        {
            $response = InvokeApiNoParseContent -Method POST -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            $statusUrl = $response.Headers['Location']

            if ($response.StatusCode -ne 200 -and $response.StatusCode -ne 202)
            {
                #Write-Error "An error occured."
                CreateHttpResponse($response)
            }
            else
            {
                $currentTime = Get-Date -format HH:mm:ss
                $nextTime = Get-Date -format HH:mm:ss
                $TimeDiff = New-TimeSpan $currentTime $nextTime
                $timeoutInSeconds = 300

                #Wait until CDS provision is completed, or we hit a timeout
                while(($response.StatusCode -eq 202) -and ($TimeDiff.TotalSeconds -lt $timeoutInSeconds))
                {
                    Start-Sleep -s 5
                    $response = InvokeApiNoParseContent -Route $statusUrl -Method GET -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
                    $nextTime = Get-Date -format HH:mm:ss
                    $TimeDiff = New-TimeSpan $currentTime $nextTime
                }

                $parsedResponse = ConvertFrom-Json $response.Content
                CreateEnvironmentObject -EnvObject $parsedResponse -ReturnCdsDatabaseType $true
            }
        }
        # optionally the caller can choose to NOT wait until provisioning is complete and get the provisioning status by polling on Get-AdminPowerAppEnvironment and looking at the provisioning status field
        else
        {
            $response = InvokeApi -Method POST -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            if ($response.StatusCode -ne 200 -and $response.StatusCode -ne 202)
            {
                #Write-Error "An error occured."
                CreateHttpResponse($response)
            }
            else
            {
                CreateEnvironmentObject -EnvObject $response -ReturnCdsDatabaseType $false
            }
        }
    }
}


function Remove-LegacyCDSDatabase
{
    <#
    .SYNOPSIS
    This command is used for removing legacy (version 1.0) CDS database.
    .DESCRIPTION
    To make sure the database exist, the command tries to get the database first. Then delete the database and verify the database is deleted.
    .PARAMETER EnvironmentName
    This is the environment name (Guid) which has CDS 1.0 database to be removed.
    .PARAMETER DatabaseId
    This is the database id for the CDS 1.0 database to be removed.
    .EXAMPLE
    Remove-LegacyCDSDatabase -EnvironmentName '0e075c48-a792-4705-8f99-82eec3b1cd8e' -DatabaseId '80c75532-5cf7-4d27-b5ef-6ab9f9024ab6'
    #>
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory=$true)]
        [string]$DatabaseId
    )
    process
    {
        $deleteDatabaseUrl = "https://api.cds.microsoft.com/providers/Microsoft.CommonDataModel/namespaces/$DatabaseId`?api-version=2016-11-01&`$filter=environment eq '$EnvironmentName'"

        # Check if the database exist
        $getDatabaseRequest = Invoke-Request -Uri $deleteDatabaseUrl -Method "GET" -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        if ($getDatabaseRequest.StatusCode -eq 200)
        {
            $deleteDatabaseRequest = Invoke-Request -Uri $deleteDatabaseUrl -Method Delete -ThrowOnFailure -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            $getDatabaseRequest = Invoke-Request -Uri $deleteDatabaseUrl -Method "GET" -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            while ($getDatabaseRequest.StatusCode -ne 404)
            {
                Start-Sleep -s 3
                Write-Verbose "Waiting for operation to complete..."

                $getDatabaseRequest = Invoke-Request -Uri $deleteDatabaseUrl -Method "GET" -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }

            Write-Verbose "Successfully removed database $DatabaseId"
        }
        elseif ($getDatabaseRequest.StatusCode -eq 404)
        {
            $deleteDatabaseRequest = $getDatabaseRequest
            Write-Verbose "Database $DatabaseId does not exist"
        }
        else
        {
            $deleteDatabaseRequest = $getDatabaseRequest
            Write-Verbose "Get database $DatabaseId fails"
        }

        return $deleteDatabaseRequest
    }
}

function Get-AdminPowerAppEnvironment
{
    <#
    .SYNOPSIS
    Returns information about one or more PowerApps environments where the calling uses is an Environment Admin. If the calling user is a tenant admin, all envionments within the tenant will be returned.
    .DESCRIPTION
    The Get-AdminPowerAppEnvironment cmdlet looks up information about =one or more environments depending on parameters.
    Use Get-Help Get-AdminPowerAppEnvironment -Examples for more detail.
    .PARAMETER Filter
    Finds environments matching the specified filter (wildcards supported).
    .PARAMETER EnvironmentName
    Finds a specific environment.
    .PARAMETER Default
    Finds the default environment.
    .PARAMETER CreatedBy
    Limit environments returned to only those created by the specified user (you can specify a email address or object id)
    .PARAMETER ApiVersion
    The api version to call with. Default 2018-01-01.
    .PARAMETER ReturnCdsDatabaseType
    If it is true, CdsDatabaseType is returned.
    .PARAMETER GetProtectedEnvironment
    Get protected environments if it is defined.
    .PARAMETER InstanceId
    Instance Id.
    .PARAMETER EnvironmentSku
    Environment Sku.
    .PARAMETER Capacity
    If it is true, capacity information is returned.
    .PARAMETER GetGenerativeAiSettings
    If it is true, generative AI settings are returned.
    .PARAMETER GetMakerAnalyticsSettings
    If it is true, maker analytics settings are returned.
    .EXAMPLE
    Get-AdminPowerAppEnvironment
    Finds all environments within the tenant where the user is an Environment Admin.
    .EXAMPLE
    Get-AdminPowerAppEnvironment -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    Finds environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    .EXAMPLE
    Get-AdminPowerAppEnvironment *Test*
    Finds all environments that contain the string "Test" in their display name where the user is an Environment Admin.
    .EXAMPLE
    Get-AdminPowerAppEnvironment -CreatedBy 7557f390-5f70-4c93-8bc4-8c2faabd2ca0
    Finds all environments that were created by the user with an object of 7557f390-5f70-4c93-8bc4-8c2faabd2ca0
    #>
    [CmdletBinding(DefaultParameterSetName="Filter")]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "User")]
        [string[]]$Filter,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Switch]$Default,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$CreatedBy,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2021-04-01",

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [bool]$ReturnCdsDatabaseType = $true,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [switch]$GetProtectedEnvironment,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$InstanceId,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$EnvironmentSku,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [Switch]$Capacity,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [Switch]$GetGenerativeAiSettings,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [Switch]$GetMakerAnalyticsSettings
    )
    process
    {
        if ($Default)
        {
            $EnvironmentName = "~default"
        }

        $expandParameter = "expand=permissions"
        if ($GetGenerativeAiSettings)
        {
            $expandParameter += ",properties/copilotPolicies"
        }

        if ($GetMakerAnalyticsSettings)
        {
            $expandParameter += ",properties.enableAnalyticsForMakers"
            $expandParameter += ",properties.enableEUPIForMakerAnalytics"
        }

        if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $getEnvironmentUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}?`$$expandParameter&api-version={apiVersion}" `
                | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

            $environmentResult = InvokeApi -Method GET -Route $getEnvironmentUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            CreateEnvironmentObject -EnvObject $environmentResult -ReturnCdsDatabaseType $ReturnCdsDatabaseType -GetGenerativeAiSettings $GetGenerativeAiSettings -GetMakerAnalyticsSettings $GetMakerAnalyticsSettings
        }
        else
        {
            if ($Capacity)
            {
                $expandParameter += ",properties.capacity"
            }

            $getAllEnvironmentsUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?`$$expandParameter&api-version={apiVersion}"
            if ($GetProtectedEnvironment)
            {
                $getAllEnvironmentsUri += "&`$filter=properties.protectionStatus.keyManagedBy eq 'Customer'"
            }

            if (-not [string]::IsNullOrWhiteSpace($InstanceId))
            {
                if ($getAllEnvironmentsUri.Contains("filter"))
                {
                    $getAllEnvironmentsUri += " and properties/linkedEnvironmentMetadata/resourceId eq {orgId}"
                }
                else
                {
                    $getAllEnvironmentsUri += "&`$filter=properties/linkedEnvironmentMetadata/resourceId eq {orgId}"
                }

                $getAllEnvironmentsUri = $getAllEnvironmentsUri | ReplaceMacro -Macro "{orgId}" -Value $InstanceId
            }

            if (-not [string]::IsNullOrWhiteSpace($EnvironmentSku))
            {
                if ($getAllEnvironmentsUri.Contains("filter"))
                {
                    $getAllEnvironmentsUri += " and properties/environmentSku eq {sku}"
                }
                else
                {
                    $getAllEnvironmentsUri += "&`$filter=properties/environmentSku eq {sku}"
                }

                $getAllEnvironmentsUri = $getAllEnvironmentsUri | ReplaceMacro -Macro "{sku}" -Value "'$EnvironmentSku'"
            }

            $environmentsResult = InvokeApi -Method GET -Route $getAllEnvironmentsUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            Get-FilteredEnvironments -Filter $Filter -CreatedBy $CreatedBy -EnvironmentResult $environmentsResult -ReturnCdsDatabaseType $ReturnCdsDatabaseType -GetGenerativeAiSettings $GetGenerativeAiSettings -GetMakerAnalyticsSettings $GetMakerAnalyticsSettings
        }
    }
}

function Get-AdminPowerAppSoftDeletedEnvironment
{
    <#
    .SYNOPSIS
    Returns information about one or more PowerApps environments which is soft-deleted. If the calling user is a tenant admin, all soft-deleted environments within the tenant will be returned.
    .DESCRIPTION
    The Get-AdminPowerAppSoftDeletedEnvironment cmdlet looks up information about one or more soft-deleted environments depending on parameters.
    Use Get-Help Get-AdminPowerAppSoftDeletedEnvironment -Examples for more detail.
    .PARAMETER Filter
    Finds soft-deleted environments matching the specified filter (wildcards supported).
    .PARAMETER EnvironmentName
    Finds a specific soft-deleted environment.
    .EXAMPLE
    Get-AdminPowerAppSoftDeletedEnvironment
    Finds all soft-deleted environments within the tenant where the user is an Tenant Admin.
    .EXAMPLE
    Get-AdminPowerAppSoftDeletedEnvironment *Test*
    Finds all soft-deleted environments that contain the string "Test" in their display name where the user is an Tenant Admin.
    #>
    [CmdletBinding(DefaultParameterSetName="Filter")]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Filter")]
        [string[]]$Filter,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $getAllEnvironmentsUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/deletedEnvironments?`$expand=permissions&api-version={apiVersion}"

        $environmentsResult = InvokeApi -Method GET -Route $getAllEnvironmentsUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        Get-FilteredEnvironments -Filter $Filter -EnvironmentResult $environmentsResult
    }
}

function Remove-AdminPowerAppEnvironment
{
    <#
    .SYNOPSIS
    Deletes the specific environment.  This operation can take some time depending on how many resources are stored in the environment. If the operation returns witha  404 NotFound, then the environment has been successfully deleted.
    .DESCRIPTION
    Remove-AdminPowerAppEnvironment cmdlet deletes an environment.
    Use Get-Help Remove-AdminPowerAppEnvironment -Examples for more detail.
    .PARAMETER Filter
    Finds environments matching the specified filter (wildcards supported).
    .EXAMPLE
    Remove-AdminPowerAppEnvironment -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    Deletes environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 and all of the environment's resources
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,


        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2018-01-01"
    )

    process
    {
        $validateEnvironmentUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/validateDelete`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        $validateResponse = InvokeApi -Method POST -Route $validateEnvironmentUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if (-not $validateResponse.canInitiateDelete)
        {
            #Write-Host "Unable to delete this environment."
            CreateHttpResponse($validateResponse)
        }
        else {

            $deleteEnvironmentUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}`?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

            $resources = $validateResponse.Content.resourcesToBeDeleted | Group type
            #Write-Host "Deleting..."
            foreach ($type in $resources)
            {
                #Write-Host $type.Name: $type.Count
            }

            #Kick-off delete
            $deleteEnvironmentResponse = InvokeApiNoParseContent -Route $deleteEnvironmentUri -Method DELETE -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            $deleteStatusUrl = $deleteEnvironmentResponse.Headers['Location']

            #If there is no status Url then the environment likely does not exist
            if(!$deleteStatusUrl)
            {
                CreateHttpResponse($deleteEnvironmentResponse)
            }
            else
            {
                # Don't poll on delete

                # $currentTime = Get-Date -format HH:mm:ss
                # $nextTime = Get-Date -format HH:mm:ss
                # $TimeDiff = New-TimeSpan $currentTime $nextTime
                # $timeoutInSeconds = 300

                # #Wait until the environment has been deleted, there is an error, or we hit a timeout
                # while($deleteEnvironmentResponse.StatusCode -ne 404 -and $deleteEnvironmentResponse.StatusCode -ne 500 -and ($TimeDiff.TotalSeconds -lt $timeoutInSeconds))
                # {
                #     Start-Sleep -s 5
                #     $deleteEnvironmentResponse = InvokeApiNoParseContent -Route $deleteStatusUrl -Method GET -ApiVersion $ApiVersion
                #     $nextTime = Get-Date -format HH:mm:ss
                #     $TimeDiff = New-TimeSpan $currentTime $nextTime
                # }

                CreateHttpResponse($deleteEnvironmentResponse)
            }
        }
    }
}

function Recover-AdminPowerAppEnvironment
{
    <#
    .SYNOPSIS
    Recovers the specific environment. This operation can take some time depending on how many resources are stored in the environment.
    .DESCRIPTION
    Recover-AdminPowerAppEnvironment cmdlet recovers an environment.
    Use Get-Help Recover-AdminPowerAppEnvironment -Examples for more detail.
    .PARAMETER EnvironmentName
    The environment name.
    .PARAMETER WaitUntilFinished
    Wait until recovery completed. The default setting is false (not wait).
    .EXAMPLE
    Recover-AdminPowerAppEnvironment -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    Recovers environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 and all of the environment's resources
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $false,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2019-05-01"
    )

    process
    {
        $recoverEnvironmentUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/deletedEnvironments/{environmentName}/recover`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        #call recover api
        $recoverEnvironmentResponse = InvokeApiNoParseContent -Route $recoverEnvironmentUri -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        $recoverStatusUrl = $recoverEnvironmentResponse.Headers['Location']

        #If there is no status Url then the environment likely does not exist
        if($recoverEnvironmentResponse.StatusCode -eq 202 -and $recoverStatusUrl -and $WaitUntilFinished)
        {
            $currentTime = Get-Date -format HH:mm:ss
            $nextTime = Get-Date -format HH:mm:ss
            $TimeDiff = New-TimeSpan $currentTime $nextTime
            $timeoutInSeconds = 300

            # Get operation status sometimes return 404 at beginning.
            while(($recoverEnvironmentResponse.StatusCode -eq 202 -or $recoverEnvironmentResponse.StatusCode -eq 404) -and ($TimeDiff.TotalSeconds -lt $timeoutInSeconds))
            {
                Start-Sleep -s 5
                $recoverEnvironmentResponse = InvokeApiNoParseContent -Method GET -Route $recoverStatusUrl -ApiVersion "2018-01-01"
                $nextTime = Get-Date -format HH:mm:ss
                $TimeDiff = New-TimeSpan $currentTime $nextTime
            }
        }

        CreateHttpResponse($recoverEnvironmentResponse)
    }
}

function Get-AdminPowerAppEnvironmentRoleAssignment
{
    <#
    .SYNOPSIS
    Returns the environment role assignments for environments without a Common Data Service For Apps database instance.
    .DESCRIPTION
    The Get-AdminPowerAppEnvironmentRoleAssignment returns environment role assignments for environments without a Common Data Service For Apps database instance.  This includes which users are assigned as an Environment Admin or Environment Maker in each environment.
    Use Get-Help Get-AdminPowerAppEnvironmentRoleAssignment -Examples for more detail.
    .PARAMETER EnvironmentName
    Limit roles returned to those in a specified environment.
    .PARAMETER UserId
    A objectId of the user you want to filter by.
    .EXAMPLE
    Get-AdminPowerAppEnvironmentRoleAssignment
    Returns all environment role assignments across all environments. where the calling users is an Environment Admin.
    .EXAMPLE
    Get-AdminPowerAppEnvironmentRoleAssignment -UserId 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    Returns all environment role assignments across all environments (where the calling users is an Environment Admin) for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    .EXAMPLE
    Get-AdminPowerAppEnvironmentRoleAssignment -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -UserId 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    Returns all environment role assignments for the environment  3c2f7648-ad60-4871-91cb-b77d7ef3c239  for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    .EXAMPLE
    Get-AdminPowerAppEnvironmentRoleAssignment  -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    Returns all environment role assignments for the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Environment", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Environment")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$UserId,

        [Parameter(Mandatory = $false, ParameterSetName = "Environment")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {

        $environments = @();

        if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $environments += @{
                EnvironmentName = $EnvironmentName
            }
        }
        else {
            $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false
        }

        foreach($environment in $environments)
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/roleAssignments?api-version={apiVersion}" `
                | ReplaceMacro -Macro "{environmentName}" -Value $environment.EnvironmentName;

            $envRoleAssignmentResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            #Write-Host $envRoleAssignmentResult.StatusCode #Returns 'Forbidden' for CDS 2.0 orgs

            $pattern = BuildFilterPattern -Filter $UserId

            foreach ($envRole in $envRoleAssignmentResult.Value)
            {
                if ($pattern.IsMatch($envRole.properties.principal.id ) -or
                    $pattern.IsMatch($envRole.properties.principal.email))
                {
                    CreateEnvRoleAssignmentObject -EnvRoleAssignmentObj $envRole -EnvObj $environment
                }
            }
        }
    }
}

function Set-AdminPowerAppEnvironmentRoleAssignment
{
<#
 .SYNOPSIS
 Sets permissions to an environment without a Common Data Service For Apps database instance. If you make this call to an environment with a Common Data Service for Apps database instance you will get a 403 Forbidden error.
 .DESCRIPTION
 The Set-AdminPowerAppEnvironmentRoleAssignment set up permission to environment depending on parameters.
 Use Get-Help Set-AdminPowerAppEnvironmentRoleAssignment -Examples for more detail.
 .PARAMETER EnvironmentName
 The environmnet id.
 .PARAMETER RoleName
 Specifies the permission level given to the environment: Environment Admin or Environment Maker.
 .PARAMETER PrincipalType
 Specifies the type of principal this environment is being shared with; a user, a security group, the entire tenant.
 .PARAMETER PrincipalObjectId
 If this environment is being shared with a user or security group principal, this field specified the ObjectId for that principal. You can use the Get-UsersOrGroupsFromGraph API to look-up the ObjectId for a user or group in Azure Active Directory.
 .EXAMPLE
 Set-AdminPowerAppEnvironmentRoleAssignment -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -RoleName EnvironmentAdmin -PrincipalType User -PrincipalObjectId 53c0a918-ce7c-401e-98f9-1c60b3a723b3
 Assigns the Environment Admin role privileges to the the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3 in the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 .EXAMPLE
 Set-AdminPowerAppEnvironmentRoleAssignment -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -RoleName EnvironmentMaker -PrincipalType Tenant
 Assigns everyone Environment Maker role privileges in the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("EnvironmentAdmin", "EnvironmentMaker")]
        [string]$RoleName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("User", "Group", "Tenant")]
        [string]$PrincipalType,

        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$PrincipalObjectId = $null,

        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $TenantId = $Global:currentSession.tenantId

        if($PrincipalType -ne "Tenant")
        {
            $userOrGroup = Get-UsersOrGroupsFromGraph -ObjectId $PrincipalObjectId
            $PrincipalEmail = $userOrGroup.Mail
        }

        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environment}/modifyRoleAssignments`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        #Construct the body
        $requestbody = $null

        If ($PrincipalType -eq "Tenant")
        {
            $requestbody = @{
                add = @(
                    @{
                        properties = @{
                            roleDefinition = @{
                                id = "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/$EnvironmentName/roleDefinitions/$RoleName"
                            }
                            principal = @{
                                email = ""
                                id = $PrincipalObjectId
                                type = $PrincipalType
                                tenantId = $TenantId
                            }
                        }
                    }
                )
            }
        }
        else
        {
            $requestbody = @{
                add = @(
                    @{
                        properties = @{
                            roleDefinition = @{
                                id = "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/$EnvironmentName/roleDefinitions/$RoleName"
                            }
                            principal = @{
                                email = $PrincipalEmail
                                id = $PrincipalObjectId
                                type = $PrincipalType
                                tenantId = "null"
                            }
                        }
                    }
                )
            }
        }

        $result = InvokeApi -Method POST -Route $route -Body $requestbody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Remove-AdminPowerAppEnvironmentRoleAssignment
{
<#
 .SYNOPSIS
 Deletes specific role assignment of an environment.
 .DESCRIPTION
 Deletes specific role assignment of an environment.
 Use Get-Help Remove-AdminPowerAppEnvironmentRoleAssignment -Examples for more detail.
 .PARAMETER EnvironmentName
 The environment id
 .PARAMETER RoleId
 Specifies the role assignment id.
 .EXAMPLE
 Remove-AdminPowerAppEnvironmentRoleAssignment -RoleId "4d1f7648-ad60-4871-91cb-b77d7ef3c239" -EnvironmentName "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877"
 Deletes the role named 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in Environment Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 #>
    [CmdletBinding(DefaultParameterSetName="App")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$RoleId,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environment}/modifyRoleAssignments`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

            #Construct the body
        $requestbody = $null

        $requestbody = @{
            remove = @(
                @{
                    id = $RoleId
                }
            )
        }


        $result = InvokeApi -Method POST -Route $route -Body $requestbody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Get-AdminPowerAppConnection
{
 <#
 .SYNOPSIS
 Returns information about one or more connection.
 .DESCRIPTION
 The Get-AdminPowerAppConnection looks up information about one or more connections depending on parameters.
 Use Get-Help Get-AdminPowerAppConnection -Examples for more detail.
 .PARAMETER Filter
 Finds connection matching the specified filter (wildcards supported).
 .PARAMETER ConnectorName
 Limit connections returned to those of a specified connector.
 .PARAMETER EnvironmentName
 Limit connections returned to those in a specified environment.
 .PARAMETER CreatedBy
 Limit connections returned to those created by by the specified user (email or AAD object id)
 .EXAMPLE
 Get-AdminPowerAppConnection
 Returns all connection from all environments where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
 .EXAMPLE
 Get-AdminPowerAppConnection *PowerApps*
 Returns all connection with the text "PowerApps" in the display namefrom all environments where the calling user is an Environment Admin  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminPowerAppConnection -CreatedBy foo@bar.onmicrosoft.com
 Returns all apps created by the user with an email of "foo@bar.onmicrosoft.com" from all environment where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminPowerAppConnection -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds connections within the 3c2f7648-ad60-4871-91cb-b77d7ef3c239 environment
  .EXAMPLE
 Get-AdminPowerAppConnection -ConnectorName shared_runtimeservice
 Finds all connections created against the shared_runtimeservice (CDS) connector from all environments where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminPowerAppConnection -ConnectorName shared_runtimeservice -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds connections within the 3c2f7648-ad60-4871-91cb-b77d7ef3c239 environment that are created against the shared_runtimeservice (CDS) connector.
 .EXAMPLE
 Get-AdminPowerAppConnection *Foobar* -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds all connections in environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 that contain the string "Foobar" in their display name.
 #>
    [CmdletBinding(DefaultParameterSetName="Filter")]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "User")]
        [string[]]$Filter,

        [Parameter(Mandatory = $false, ParameterSetName = "Connector", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Connector", ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$CreatedBy,

        [Parameter(Mandatory = $false, ParameterSetName = "Connector")]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        # If the connector name is specified, only return connections for that connector
        if (-not [string]::IsNullOrWhiteSpace($ConnectorName))
        {
            $environments = @();

            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $environments += @{
                    EnvironmentName = $EnvironmentName
                }
            }
            else
            {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false
            }

            foreach($environment in $environments)
            {
                $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connectorName}/connections?api-version={apiVersion}" `
                | ReplaceMacro -Macro "{connectorName}" -Value $ConnectorName `
                | ReplaceMacro -Macro "{environment}" -Value $environment.EnvironmentName;

                $connectionResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                Get-FilteredConnections -Filter $Filter -CreatedBy $CreatedBy -ConnectionResult $connectionResult
            }
        }
        else
        {
            # If the caller passed in an environment scope, filter the query to only that environment
            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/connections?api-version={apiVersion}" `
                | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

                $connectionResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                Get-FilteredConnections -Filter $Filter -CreatedBy $CreatedBy -ConnectionResult $connectionResult
            }
            # otherwise search for the apps acroos all environments for this calling user
            else
            {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false

                foreach($environment in $environments)
                {
                    $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/connections?api-version={apiVersion}" `
                    | ReplaceMacro -Macro "{environment}" -Value $environment.EnvironmentName;

                    $connectionResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                    Get-FilteredConnections -Filter $Filter -CreatedBy $CreatedBy -ConnectionResult $connectionResult
                }
            }
        }
    }
}


function Remove-AdminPowerAppConnection
{
 <#
 .SYNOPSIS
 Deletes the connection.
 .DESCRIPTION
 The Remove-AdminPowerAppConnection permanently deletes the connection.
 Use Get-Help Remove-AdminPowerAppConnection -Examples for more detail.
 .PARAMETER ConnectionName
 The connection identifier.
 .PARAMETER ConnectorName
 The connection's connector name.
 .PARAMETER EnvironmentName
 The connection's environment.
 .EXAMPLE
 Remove-AdminPowerAppConnection -ConnectionName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -ConnectorName shared_twitter -EnvironmentName Default-efecdc9a-c859-42fd-b215-dc9c314594dd
 Deletes the connection with name 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$ConnectionName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$ConnectorName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}/connections/{connection}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
        | ReplaceMacro -Macro "{connection}" -Value $ConnectionName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $removeResult = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($removeResult -eq $null)
        {
            return $null
        }

        CreateHttpResponse($removeResult)
    }
}

function Get-AdminPowerAppConnectionRoleAssignment
{
 <#
 .SYNOPSIS
 Returns the connection role assignments for a user or a connection. Owner role assignments cannot be deleted without deleting the connection resource.
 .DESCRIPTION
 The Get-AdminPowerAppConnectionRoleAssignment functions returns all roles assignments for an connection or all connection roles assignments for a user (across all of their connections).  A connection's role assignments determine which users have access to the connection for using or building apps and flows and with which permission level (CanUse, CanUseAndShare) .
 Use Get-Help Get-AdminPowerAppConnectionRoleAssignment -Examples for more detail.
 .PARAMETER ConnectionName
 The connection identifier.
 .PARAMETER EnvironmentName
 The connections's environment.
 .PARAMETER ConnectorName
 The connection's connector identifier.
 .PARAMETER PrincipalObjectId
 The objectId of a user or group, if specified, this function will only return role assignments for that user or group.
 .EXAMPLE
 Get-AdminPowerAppConnectionRoleAssignment
 Returns all connection role assignments for the calling user.
 .EXAMPLE
 Get-AdminPowerAppConnectionRoleAssignment -ConnectionName 3b4b9592607147258a4f2fb33517e97a -ConnectorName shared_sharepointonline -EnvironmentName ee1eef10-ba55-440b-a009-ce379f86e20c
 Returns all role assignments for the connection with name 3b4b9592607147258a4f2fb33517e97ain environment with name ee1eef10-ba55-440b-a009-ce379f86e20c for the connector named shared_sharepointonline
 .EXAMPLE
 Get-AdminPowerAppConnectionRoleAssignment -ConnectionName 3b4b9592607147258a4f2fb33517e97a -ConnectorName shared_sharepointonline -EnvironmentName ee1eef10-ba55-440b-a009-ce379f86e20c -PrincipalObjectId 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Returns all role assignments for the user, or group with an object of 3c2f7648-ad60-4871-91cb-b77d7ef3c239 for the connection with name 3b4b9592607147258a4f2fb33517e97ain environment with name ee1eef10-ba55-440b-a009-ce379f86e20c for the connector named shared_sharepointonline
 #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectionName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$PrincipalObjectId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $selectedObjectId = $null

        if (-not [string]::IsNullOrWhiteSpace($ConnectionName))
        {
            if (-not [string]::IsNullOrWhiteSpace($PrincipalObjectId))
            {
                $selectedObjectId = $PrincipalObjectId;
            }
        }

        $pattern = BuildFilterPattern -Filter $selectedObjectId

        if (-not [string]::IsNullOrWhiteSpace($ConnectionName))
        {

            $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}/connections/{connection}/permissions?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
            | ReplaceMacro -Macro "{connection}" -Value $ConnectionName `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

            $connectionRoleResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            foreach ($connectionRole in $connectionRoleResult.Value)
            {
                if (-not [string]::IsNullOrWhiteSpace($PrincipalObjectId))
                {
                    if ($pattern.IsMatch($connectionRole.properties.principal.id ) -or
                        $pattern.IsMatch($connectionRole.properties.principal.email) -or
                        $pattern.IsMatch($connectionRole.properties.principal.tenantId))
                    {
                        CreateConnectionRoleAssignmentObject -ConnectionRoleAssignmentObj $connectionRole -EnvironmentName $EnvironmentName
                    }
                }
                else
                {
                    CreateConnectionRoleAssignmentObject -ConnectionRoleAssignmentObj $connectionRole -EnvironmentName $EnvironmentName
                }
            }
        }
        else
        {
            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $connections = Get-AdminPowerAppConnection -EnvironmentName $EnvironmentName -ConnectorName $ConnectorName -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }
            else
            {
                $connections = Get-AdminPowerAppConnection -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }

            foreach($connection in $connections)
            {
                Get-AdminPowerAppConnectionRoleAssignment `
                    -ConnectionName $connection.ConnectionName `
                    -ConnectorName $connection.ConnectorName `
                    -EnvironmentName $connection.EnvironmentName `
                    -PrincipalObjectId $selectedObjectId `
                    -ApiVersion $ApiVersion `
                    -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }
        }
    }
}

function Set-AdminPowerAppConnectionRoleAssignment
{
    <#
    .SYNOPSIS
    Sets permissions to the connection.
    .DESCRIPTION
    The Set-AdminPowerAppConnectionRoleAssignment set up permission to connection depending on parameters.
    Use Get-Help Set-AdminPowerAppConnectionRoleAssignment -Examples for more detail.
    .PARAMETER ConnectionName
    The connection identifier.
    .PARAMETER EnvironmentName
    The connections's environment.
    .PARAMETER ConnectorName
    The connection's connector identifier.
    .PARAMETER RoleName
    Specifies the permission level given to the connection: CanView, CanViewWithShare, CanEdit. Sharing with the entire tenant is only supported for CanView.
    .PARAMETER PrincipalType
    Specifies the type of principal this connection is being shared with; a user, a security group, the entire tenant.
    .PARAMETER PrincipalObjectId
    If this connection is being shared with a user or security group principal, this field specified the ObjectId for that principal. You can use the Get-UsersOrGroupsFromGraph API to look-up the ObjectId for a user or group in Azure Active Directory.
    .EXAMPLE
    Set-AdminPowerAppConnectionRoleAssignment -PrincipalType Group -PrincipalObjectId b049bf12-d56d-4b50-8176-c6560cbd35aa -RoleName CanEdit -ConnectionName 3b4b9592607147258a4f2fb33517e97a -ConnectorName shared_vsts -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
    Give the specified security group CanEdit permissions to the connection with name 3b4b9592607147258a4f2fb33517e97a
    #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectionName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("CanView", "CanViewWithShare", "CanEdit")]
        [string]$RoleName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("User", "Group", "Tenant")]
        [string]$PrincipalType,

        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$PrincipalObjectId = $null,

        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $TenantId = $Global:currentSession.tenantId

        # if($PrincipalType -ne "Tenant")
        # {
        #     $userOrGroup = Get-UsersOrGroupsFromGraph -ObjectId $PrincipalObjectId
        #     $PrincipalEmail = $userOrGroup.Mail
        # }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}/connections/{connection}/modifyPermissions?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
        | ReplaceMacro -Macro "{connection}" -Value $ConnectionName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        #Construct the body
        $requestbody = $null

        If ($PrincipalType -eq "Tenant")
        {
            $requestbody = @{
                delete = @()
                put = @(
                    @{
                        properties = @{
                            roleName = $RoleName
                            principal = @{
                                id = $TenantId
                                tenantId = $TenantId
                            }
                        }
                    }
                )
            }
        }
        else
        {
            $requestbody = @{
                delete = @()
                put = @(
                    @{
                        properties = @{
                            roleName = $RoleName
                            principal = @{
                                id = $PrincipalObjectId
                            }
                        }
                    }
                )
            }
        }

        $setConnectionRoleResult = InvokeApi -Method POST -Body $requestbody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($setConnectionRoleResult)
    }
}

function Remove-AdminPowerAppConnectionRoleAssignment
{
 <#
 .SYNOPSIS
 Deletes a connection role assignment record.
 .DESCRIPTION
 The Remove-AdminPowerAppConnectionRoleAssignment deletes the specific connection role assignment
 Use Get-Help Remove-AdminPowerAppConnectionRoleAssignment -Examples for more detail.
 .PARAMETER RoleId
 The id of the role assignment to be deleted.
 .PARAMETER ConnectionName
 The app identifier.
 .PARAMETER ConnectorName
 The connection's associated connector name
 .PARAMETER EnvironmentName
 The connection's environment.
 .EXAMPLE
 Remove-AdminPowerAppConnectionRoleAssignment -ConnectionName a2956cf95ba441119d16dc2ef0ca1ff9 -EnvironmentName 08b4e32a-4e0d-4a69-97da-e1640f0eb7b9 -ConnectorName shared_twitter -RoleId /providers/Microsoft.PowerApps/apis/shared_twitter/connections/a2956cf95ba441119d16dc2ef0ca1ff9/permissions/7557f390-5f70-4c93-8bc4-8c2faabd2ca0
 Deletes the app role assignment with an id of /providers/Microsoft.PowerApps/apps/f8d7a19d-f8f9-4e10-8e62-eb8c518a2eb4/permissions/tenant-efecdc9a-c859-42fd-b215-dc9c314594dd
 #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectionName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$RoleId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2017-06-01"
    )

    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apis/{connector}/connections/{connection}/modifyPermissions`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{connector}" -Value $ConnectorName `
        | ReplaceMacro -Macro "{connection}" -Value $ConnectionName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        #Construct the body
        $requestbody = $null

        $requestbody = @{
            delete = @(
                @{
                    id = $RoleId
                }
            )
        }

        $removeResult = InvokeApi -Method POST -Body $requestbody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($removeResult -eq $null)
        {
            return $null
        }

        CreateHttpResponse($removeResult)
    }
}


function Get-AdminPowerApp
{
 <#
 .SYNOPSIS
 Returns information about one or more apps.
 .DESCRIPTION
 The Get-AdminPowerApp looks up information about one or more apps depending on parameters.
 Use Get-Help Get-AdminPowerApp -Examples for more detail.
 .PARAMETER Filter
 Finds apps matching the specified filter (wildcards supported).
 .PARAMETER AppName
 Finds a specific id.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .PARAMETER Owner
 Limit apps returned to those owned by the specified user (you can specify a email address or object id)
 .EXAMPLE
 Get-AdminPowerApp
 Returns all apps from all environments where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
 .EXAMPLE
 Get-AdminPowerApp *PowerApps*
 Returns all apps with the text "PowerApps" from all environments where the calling user is an Environment Admin  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminPowerApp -CreatedBy foo@bar.onmicrosoft.com
 Returns all apps created by the user with an email of "foo@bar.onmicrosoft.com" from all environment where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
  .EXAMPLE
 Get-AdminPowerApp -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds apps within the 3c2f7648-ad60-4871-91cb-b77d7ef3c239 environment
 .EXAMPLE
 Get-AdminPowerApp *Foobar* -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds all app in environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 that contain the string "Foobar" in their display name.
 .EXAMPLE
 Get-AdminPowerApp -AppName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Returns the details for the app named 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in Environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 #>
    [CmdletBinding(DefaultParameterSetName="Filter")]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "User")]
        [string[]]$Filter,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$Owner,

        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        # If the app name is specified, just return the details for that app
        if (-not [string]::IsNullOrWhiteSpace($AppName))
        {
            $top = 250
            $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps/{appName}?api-version={apiVersion}&`$top={top}&`$expand=unpublishedAppDefinition" `
            | ReplaceMacro -Macro "{appName}" -Value $AppName `
            | ReplaceMacro -Macro "{top}" -Value $top `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

            $appResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            CreateAppObject -AppObj $appResult;
        }
        else
        {
            $userId = $Global:currentSession.userId
            $expandPermissions = "permissions(`$filter=maxAssignedTo(`'$userId`'))"

            # If the caller passed in an environment scope, filter the query to only that environment
            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $top = 250

                $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps?api-version={apiVersion}&`$top={top}&`$expand={expandPermissions}" `
                | ReplaceMacro -Macro "{top}" -Value $top `
                | ReplaceMacro -Macro "{expandPermissions}" -Value $expandPermissions `
                | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

                $appResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                Get-FilteredApps -Filter $Filter -Owner $Owner -AppResult $appResult
            }
            # otherwise search for the apps across all environments for this calling user
            else
            {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false

                foreach($environment in $environments)
                {
                    $top = 250

                    $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps?api-version={apiVersion}&`$top={top}&`$expand={expandPermissions}" `
                    | ReplaceMacro -Macro "{top}" -Value $top `
                    | ReplaceMacro -Macro "{expandPermissions}" -Value $expandPermissions `
                    | ReplaceMacro -Macro "{environment}" -Value $environment.EnvironmentName;

                    $appResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                    Get-FilteredApps -Filter $Filter -Owner $Owner -AppResult $appResult
                }
            }
        }
    }
}

function Remove-AdminPowerApp
{
<#
 .SYNOPSIS
 Deletes an app.
 .DESCRIPTION
 The Remove-AdminPowerApp deletes an app.
 Use Get-Help Remove-AdminPowerApp -Examples for more detail.
 .PARAMETER AppName
 Specifies the app id.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Remove-AdminPowerApp -AppName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Deletes the app named 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in Environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 #>
    [CmdletBinding(DefaultParameterSetName="App")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps/{appName}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $result = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Get-AdminPowerAppRoleAssignment
{
<#
 .SYNOPSIS
 Returns permission information about one or more apps.
 .DESCRIPTION
 The Get-AdminPowerAppRoleAssignment returns permission information about one or more apps.
 Use Get-Help Get-AdminPowerAppRoleAssignment -Examples for more detail.
 .PARAMETER AppName
 Finds a specific id.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .PARAMETER UserId
 A objectId of the user you want to filter by.
 .EXAMPLE
 Get-AdminPowerAppRoleAssignments -UserId 53c0a918-ce7c-401e-98f9-1c60b3a723b3
 Returns all app role assignments across all environments for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
 .EXAMPLE
 Get-AdminPowerAppRoleAssignments -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -UserId 53c0a918-ce7c-401e-98f9-1c60b3a723b3
 Returns all app role assignemtns within environment with id 3c2f7648-ad60-4871-91cb-b77d7ef3c239 for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
 .EXAMPLE
 Get-AdminPowerAppRoleAssignments -AppName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -UserId 53c0a918-ce7c-401e-98f9-1c60b3a723b3
 Returns all role assignments for the app with id 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
 .EXAMPLE
 Get-AdminPowerAppRoleAssignments -AppName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Returns all role assignments for the app with id 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "Environment", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [Parameter(Mandatory = $true, ParameterSetName = "Environment")]
        [string]$UserId,

        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [Parameter(Mandatory = $false, ParameterSetName = "Environment")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        if (-not [string]::IsNullOrWhiteSpace($AppName))
        {
            $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps/{appName}/permissions?api-version={apiVersion}&`$filter=environment%20eq%20%27{environment}%27" `
            | ReplaceMacro -Macro "{appName}" -Value $AppName `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

            $appRoleAssignmentResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            $pattern = BuildFilterPattern -Filter $UserId

            foreach ($appRole in $appRoleAssignmentResult.Value)
            {
                if ($pattern.IsMatch($appRole.properties.principal.id ) -or
                    $pattern.IsMatch($appRole.properties.principal.email))
                {
                    CreateAppRoleAssignmentObject -AppRoleAssignmentObj $appRole
                }
            }
        }
        else
        {
            $environments = @();

            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $environments += @{
                    EnvironmentName = $EnvironmentName
                }
            }
            else {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false
            }

            foreach($environment in $environments)
            {
                $appResult = Get-AdminPowerApp -EnvironmentName $environment.EnvironmentName

                foreach ($app in $appResult)
                {
                    $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps/{appName}/permissions?api-version={apiVersion}&`$filter=environment%20eq%20%27{environment}%27" `
                    | ReplaceMacro -Macro "{appName}" -Value $app.AppName `
                    | ReplaceMacro -Macro "{environment}" -Value $environment.EnvironmentName;

                    $appRoleAssignmentResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                    $pattern = BuildFilterPattern -Filter $UserId

                    foreach ($appRole in $appRoleAssignmentResult.Value)
                    {
                        if ($pattern.IsMatch($appRole.properties.principal.id ) -or
                            $pattern.IsMatch($appRole.properties.principal.email))
                        {
                            CreateAppRoleAssignmentObject -AppRoleAssignmentObj $appRole
                        }
                    }
                }
            }
        }
    }
}

function Add-PowerAppsCustomBrandingAssets
{
    <#
    .SYNOPSIS
    Uploads tenant level custom assets. Please note that
    1. The 'mobile' parameter is required in the request body.
    2. The 'tenantSplashImagePath' and 'tenantWelcomeImagePath' parameters should contain the file paths to the image you would like to upload
    3. The tenantSplashImage cannot exceed dimensions 384x384 and the tenantWelcomeImage cannot
    exceed dimensions 1335x1491.
    4. the tenantSplashImage file size cannot exceed 3MB and the tenantWelcomeImage cannot exceed 5MB
    .DESCRIPTION
    The Add-PowerAppsCustomBrandingAssets function uploads custom assets for a tenant.
    These assets will enable customization of icons and color preferences on opening of the
    Power Apps mobile app.
    Use Get-Help Add-PowerAppsCustomBrandingAssets -Examples for more detail.
    .PARAMETER RequestBody
    The JSON object containing the custom asset info parameters. Please see examples for an example of the parameters.
    .EXAMPLE
    Add-PowerAppsCustomBrandingAssets -AssetInfo @{
        "mobile"= "default";
        "fillColor"="#CCCCCC";
        "actionButtonHighlight"= "#C43E1C";
        "hyperlinkColor"= "#FF6709";
        "headingTextColor"= "#FFFFFF";
        "statusBarContentColorMode"= "Dark";
        "tenantSplashImagePath"= "./Resources/image1.png";
        "tenantWelcomeImagePath"= "./Resources/image2.png";
        "defaultAppIds"= @("2b9a6e99-7722-4cba-bbdd-190748b47a66")
    }
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$RequestBody,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $mobileType = $RequestBody.mobile

        if ([string]::IsNullOrEmpty($mobileType))
        {
            Write-Error -Message "Mobile type must not be null or empty." -ErrorAction Stop
        }

        $iconImagePath = $RequestBody.tenantSplashImagePath
        $welcomeImagePath = $RequestBody.tenantWelcomeImagePath
        $containerUri = $null
        $sas = $null
        $tenantSplashImageUri = $null
        $tenantWelcomeImageUri = $null
        if (![string]::IsNullOrEmpty($iconImagePath) -or ![string]::IsNullOrEmpty($welcomeImagePath)) {
            $generateStorageRoute = "https://{powerAppsEndpoint}/generateTenantResourceStorage?api-version=$ApiVersion"

            try {
                $containerResult = InvokeApi -Method POST -Route $generateStorageRoute -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
                $containerSharedAccessSignature = $containerResult.sharedAccessSignature
                $containerUri = $containerSharedAccessSignature.split('?')[0]
                $sas = $containerSharedAccessSignature.split('?')[1]

                Write-Verbose "Successfully generated tenant level container"
            } catch {
                $response = $_.Exception.Response
                if ($_.ErrorDetails)
                {
                    $errorResponse = ConvertFrom-Json $_.ErrorDetails;
                    $code = $response.StatusCode
                    $message = $errorResponse.error.message

                    Write-Verbose "Status Code: '$code'. Message: '$message'"
                }

                Write-Error -Message "Error encountered while generating container" -ErrorAction Stop
            }
        }

        $headers = @{ 'x-ms-blob-type' = 'BlockBlob' }
        if (![string]::IsNullOrEmpty($iconImagePath))
        {
            $tenantSplashImageUri = $containerUri + "/tenantSplashImage?" + $sas

            Write-Verbose "tenantSplashImageUri:  '$tenantSplashImageUri'"
            Write-Verbose "iconImagePath: '$iconImagePath'"

            try {
                Invoke-RestMethod -Uri $tenantSplashImageUri -Method Put -Headers $headers -InFile $iconImagePath

                Write-Verbose "Successfully saved images to container"
            } catch {
                $response = $_.Exception.Response
                if ($_.ErrorDetails)
                {
                    $errorResponse = ConvertFrom-Json $_.ErrorDetails;
                    $code = $response.StatusCode
                    $message = $errorResponse.error.message

                    Write-Verbose "Status Code: '$code'. Message: '$message'"
                }

                Write-Error -Message "Error encountered while saving image to container" -ErrorAction Stop
            }

        }

        if (![string]::IsNullOrEmpty($welcomeImagePath))
        {
            $tenantWelcomeImageUri = $containerUri + "/tenantWelcomeImage?" + $sas

            try {
                $welcomeImageResponse = Invoke-RestMethod -Uri $tenantWelcomeImageUri -Method Put -Headers $headers -InFile $welcomeImagePath

                Write-Verbose "Successfully saved images to container"
            } catch {
                $response = $_.Exception.Response
                if ($_.ErrorDetails)
                {
                    $errorResponse = ConvertFrom-Json $_.ErrorDetails;
                    $code = $response.StatusCode
                    $message = $errorResponse.error.message

                    Write-Verbose "Status Code: '$code'. Message: '$message'"
                }

                Write-Error -Message "Error encountered while saving image to container" -ErrorAction Stop
            }
        }

        $customAssetsRequestBody = @{
            'fillColor'= $RequestBody.fillColor
            'buttonColor'= $RequestBody.buttonColor
            'actionButtonHighlight'= $RequestBody.actionButtonHighlight
            'hyperlinkColor'= $RequestBody.hyperlinkColor
            'headingTextColor'= $RequestBody.headingTextColor
            'statusBarContentColorMode'= $RequestBody.statusBarContentColorMode
            'customAssetsLinks'= @{
                'tenantSplashImage'= $tenantSplashImageUri
                'tenantWelcomeImage'= $tenantWelcomeImageUri
            }
            'defaultAppIds'=$RequestBody.defaultAppIds
        }

        $uploadRoute = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/customAssets/" + $mobileType + "?api-version=$ApiVersion"

        $uploadResult = InvokeApi -Method POST -Route $uploadRoute -Body $customAssetsRequestBody -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($uploadResult)
    }
}

function Remove-AdminPowerAppRoleAssignment
{
<#
 .SYNOPSIS
 Deletes specific role of an app.
 .DESCRIPTION
 Deletes specific role of an app.
 Use Get-Help Remove-AdminPowerAppRoleAssignment -Examples for more detail.
 .PARAMETER AppName
 Specifies the app id.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .PARAMETER RoleId
 Specifies the role assignment id.
 Remove-AdminPowerAppRoleAssignment -RoleId "4d1f7648-ad60-4871-91cb-b77d7ef3c239" -EnvironmentName "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" -AppName "73691d1f-0ff5-442c-87ce-1e3e2fba58dc"
 Deletes the role named 4d1f7648-ad60-4871-91cb-b77d7ef3c239 for app id 73691d1f-0ff5-442c-87ce-1e3e2fba58dc in Environment Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 #>
    [CmdletBinding(DefaultParameterSetName="App")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$RoleId,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps/{appName}/modifyPermissions?api-version={apiVersion}&`$filter=environment%20eq%20%27{environment}%27" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

            #Construct the body
        $requestbody = $null

        $requestbody = @{
            delete = @(
                @{
                    id = $RoleId
                 }
             )
             }

        $result = InvokeApi -Method POST -Route $route -Body $requestbody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Set-AdminPowerAppRoleAssignment
{
<#
 .SYNOPSIS
 sets permissions to the app.
 .DESCRIPTION
 The Set-AdminPowerAppRoleAssignment set up permission to app depending on parameters.
 Use Get-Help Set-AdminPowerAppRoleAssignment -Examples for more detail.
 .PARAMETER AppName
 App name for the one which you want to set permission.
 .PARAMETER EnvironmentName
 Limit app returned to those in a specified environment.
 .PARAMETER RoleName
 Specifies the permission level given to the app: CanView, CanEdit. Sharing with the entire tenant is only supported for CanView.
 .PARAMETER PrincipalType
 Specifies the type of principal this app is being shared with; a user, a security group, the entire tenant.
 .PARAMETER PrincipalObjectId
 If this app is being shared with a user or security group principal, this field specified the ObjectId for that principal. You can use the Get-UsersOrGroupsFromGraph API to look-up the ObjectId for a user or group in Azure Active Directory.
 .PARAMETER Notify
 Specifies the option (Notify, DoNotNotify) on notifying the share target.
 .EXAMPLE
 Set-AdminPowerAppRoleAssignment -PrincipalType Group -PrincipalObjectId b049bf12-d56d-4b50-8176-c6560cbd35aa -RoleName CanEdit -AppName 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488 -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 Give the specified security group CanEdit permissions to the app with name 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488
 #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [ValidateSet("CanView", "CanEdit")]
        [string]$RoleName,

        [Parameter(Mandatory = $true, ParameterSetName = "Tenant")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [ValidateSet("User", "Group", "Tenant")]
        [string]$PrincipalType,

        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$PrincipalObjectId = $null,

        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [string]$ApiVersion = "2016-11-01",

        [Parameter(Mandatory = $false, ParameterSetName = "Tenant")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [ValidateSet("Notify", "DoNotNotify")]
        [string]$Notify = "Notify"
    )

    process
    {
        $TenantId = $Global:currentSession.tenantId

        if($PrincipalType -ne "Tenant")
        {
            $userOrGroup = Get-UsersOrGroupsFromGraph -ObjectId $PrincipalObjectId
            $PrincipalEmail = $userOrGroup.Mail
        }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps/{appName}/modifyPermissions`?api-version={apiVersion}&`$filter=environment%20eq%20'{environment}'" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        #Construct the body
        $requestbody = $null

        If ($PrincipalType -eq "Tenant")
        {
            $requestbody = @{
                put = @(
                    @{
                        properties = @{
                            roleName = $RoleName
                            capabilities = @()
                            NotifyShareTargetOption = $Notify
                            principal = @{
                                type = $PrincipalType
                                tenantId = $TenantId
                            }
                        }
                    }
                )
            }
        }
        else
        {
            $requestbody = @{
                put = @(
                    @{
                        properties = @{
                            roleName = $RoleName
                            capabilities = @()
                            NotifyShareTargetOption = $Notify
                            principal = @{
                                email = $PrincipalEmail
                                id = $PrincipalObjectId
                                type = $PrincipalType
                                tenantId = "null"
                            }
                        }
                    }
                )
            }
        }

        $result = InvokeApi -Method POST -Route $route -Body $requestbody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Set-AdminPowerAppOwner
{
<#
 .SYNOPSIS
 Sets the app owner and changes the current owner to "Can View" role type.
 .DESCRIPTION
 The Set-AppOwner Sets the app owner and changes the current owner to "Can View" role type.
 Use Get-Help Set-AppOwner -Examples for more detail.
 .PARAMETER AppName
 App name for the one which you want to set permission.
 .PARAMETER EnvironmentName
 Limit app returned to those in a specified environment.
 .PARAMETER AppOwner
 Id of new owner which you want to set.
 .EXAMPLE
 Set-AppOwner -AppName "73691d1f-0ff5-442c-87ce-1e3e2fba58dc" -EnvironmentName "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" -AppOwner "1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488"
 Sets the app owner to "1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488" and changes the current owner to "Can View" role type.
 #>
    [CmdletBinding(DefaultParameterSetName="App")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [string]$ApiVersion = "2016-11-01",

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "App")]
        [string]$AppOwner
    )

    process
    {

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/apps/{appName}/modifyAppOwner?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{appName}" -Value $AppName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        #Construct the body
        $requestbody = $null

        $requestbody =
                @{
                    newAppOwner = $AppOwner
                    roleForOldAppOwner = "CanView"
                    }

        $result = InvokeApi -Method POST -Route $route -Body $requestbody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)

    }
}

function Get-AdminFlow
{
<#
 .SYNOPSIS
 Returns information about one or more flows.
 .DESCRIPTION
 The Get-AdminFlow looks up information about one or more flows depending on parameters.
 Use Get-Help Get-AdminFlow -Examples for more detail.
 .PARAMETER Filter
 Finds flows matching the specified filter (wildcards supported).
 .PARAMETER FlowName
 Finds a specific ID.
 .PARAMETER EnvironmentName
 Limit flows returned to those in a specified environment.
 .PARAMETER CreatedBy
 Limit flows returned to those created by the specified user (you must specify a user's object ID).
 .PARAMETER IncludeDeleted
 Include soft-deleted flows in the returned list.
 .PARAMETER IncludeEUDBNonCompliantFlows
 Include non-EUDB compliant flows in the returned list
  .EXAMPLE
 Get-AdminFlow
 Returns all flow from all environments where the current user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
 .EXAMPLE
 Get-AdminFlow -CreatedBy dbfad833-1e1e-4665-a20c-96391a1a39f0
 Returns all apps created by the user with an object of "dbfad833-1e1e-4665-a20c-96391a1a39f0" from all environment where the calling user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
 .EXAMPLE
 Get-AdminFlow *Flows*
 Returns all flows with the text "Flows" from all environments where the current user is an Environment Admin.  For Global admins, this will search across all environments in the tenant.
 .EXAMPLE
 Get-AdminFlow -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds flows within the 3c2f7648-ad60-4871-91cb-b77d7ef3c239 environment
 .EXAMPLE
 Get-AdminFlow -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -IncludeDeleted $true
 Finds flows within the 3c2f7648-ad60-4871-91cb-b77d7ef3c239 environment, including any that are soft-deleted
 .EXAMPLE
 Get-AdminFlow *Foobar* -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds all flows in environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 that contain the string "Foobar" in their display name.
 .EXAMPLE
 Get-AdminFlow -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Returns the details for the flow named 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in Environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 .EXAMPLE
 Get-AdminFlow -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -IncludeEUDBNonCompliantFlows $true
 Finds flows that are not EUDB compliant within the 3c2f7648-ad60-4871-91cb-b77d7ef3c239 environment.
 #>
    [CmdletBinding(DefaultParameterSetName="Filter")]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Filter")]
        [string[]]$Filter,

        [Parameter(Mandatory = $false, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "App", ValueFromPipelineByPropertyName = $true)]
        [string]$FlowName,

        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$CreatedBy,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [boolean]$IncludeDeleted,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [boolean]$IncludeEUDBNonCompliantFlows,

        [Parameter(Mandatory = $false, ParameterSetName = "App")]
        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $includeParam = if ($IncludeDeleted -and $IncludeEUDBNonCompliantFlows)
        {
            "&include=softDeletedFlows,flowsInUserResourceGroup"
        }
        elseif ($IncludeDeleted)
        {
            "&include=softDeletedFlows"
        }
        elseif ($IncludeEUDBNonCompliantFlows)
        {
            "&include=flowsInUserResourceGroup"
        }
        else {
            ""
        }

        if (-not [string]::IsNullOrWhiteSpace($FlowName))
        {
            $top = 50
            $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}?api-version={apiVersion}&`$top={top}$($includeParam)" `
            | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
            | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName) `
            | ReplaceMacro -Macro "{top}" -Value $top;

            $flowResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            CreateFlowObject -FlowObj $flowResult;
        }
        else
        {
            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $top = 50

                $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/v2/flows?api-version={apiVersion}&`$top={top}$($includeParam)" `
                | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
                | ReplaceMacro -Macro "{top}" -Value $top;

                $flowResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                Get-FilteredFlows -Filter $Filter -CreatedBy $CreatedBy -FlowResult $flowResult
            }
            else {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false

                foreach($environment in $environments)
                {
                    $top = 50

                    $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/v2/flows?api-version={apiVersion}&`$top={top}$($includeParam)" `
                    | ReplaceMacro -Macro "{environment}" -Value $environment.EnvironmentName `
                    | ReplaceMacro -Macro "{top}" -Value $top;

                    $flowResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                    Get-FilteredFlows -Filter $Filter -CreatedBy $CreatedBy -FlowResult $flowResult
                }
            }
        }
    }
}


function Enable-AdminFlow
{
<#
 .SYNOPSIS
 Starts the specific flow.
 .DESCRIPTION
 The Enable-AdminFlow starts the specific flow.
 Use Get-Help Enable-AdminFlow -Examples for more detail.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Enable-AdminFlow -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239
 Starts the 4d1f7648-ad60-4871-91cb-b77d7ef3c239 flow in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877"
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/start?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $result = InvokeApi -Method POST -Route $route -Body @{} -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Disable-AdminFlow
{
<#
 .SYNOPSIS
 Stops the specific flow.
 .DESCRIPTION
 The Disable-AdminFlow stops the specific flow.
 Use Get-Help Disable-AdminFlow -Examples for more detail.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Disable-AdminFlow -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239
 Stops the 4d1f7648-ad60-4871-91cb-b77d7ef3c239 flow in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877"
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,


        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/stop?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $result = InvokeApi -Method POST -Route $route -Body @{} -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Remove-AdminFlow
{
<#
 .SYNOPSIS
 Delete the specific flow.
 .DESCRIPTION
 The Remove-AdminFlow deletes the specific flow.
 Use Get-Help Remove-AdminFlow -Examples for more detail.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Remove-AdminFlow -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239
 Deletes the 4d1f7648-ad60-4871-91cb-b77d7ef3c239 flow in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877"
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $result = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Restore-AdminFlow
{
<#
 .SYNOPSIS
 Restore the specified soft-deleted flow.
 .DESCRIPTION
 The Restore-AdminFlow command deletes the specific flow.
 Use Get-Help Restore-AdminFlow -Examples for more detail.
 .PARAMETER FlowName
 Specifies the flow ID.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Restore-AdminFlow -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239
 Restores the soft-deleted 4d1f7648-ad60-4871-91cb-b77d7ef3c239 flow in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877"
 .EXAMPLE
 foreach ($id in @(
     "4d1f7648-ad60-4871-91cb-b77d7ef3c239",
     "eb2266a8-67b6-4919-8afd-f59c3c0e4131"
 )) {
   Restore-AdminFlow -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName $id
   Start-Sleep -Seconds 1
 }
 Restores the soft-deleted 4d1f7648-ad60-4871-91cb-b77d7ef3c239 and eb2266a8-67b6-4919-8afd-f59c3c0e4131 flows in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877"
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/restore?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $result = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Set-AdminFlowOwnerRole
{
<#
 .SYNOPSIS
 sets owner permissions to the flow.
 .DESCRIPTION
 The Set-AdminFlowOwnerRole set up permission to flow depending on parameters.
 Use Get-Help Set-AdminFlowOwnerRole -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit app returned to those in a specified environment.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER RoleName
 Specifies the access level for the user on the flow; CanView or CanEdit
 .PARAMETER PrincipalType
 Specifies the type of principal that is being added as an owner; User or Group (security group)
 .PARAMETER PrincipalObjectId
 Specifies the principal object Id of the user or security group.
 .EXAMPLE
 Set-AdminFlowOwnerRole -PrincipalType Group -PrincipalObjectId b049bf12-d56d-4b50-8176-c6560cbd35aa -RoleName CanEdit -FlowName 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488 -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 Add the specified security group as an owner fo the flow with name 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488
 #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$FlowName,

        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [ValidateSet("User", "Group")]
        [string]$PrincipalType,

        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [ValidateSet("CanView", "CanEdit")]
        [string]$RoleName,

        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [string]$PrincipalObjectId = $null,

        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $userOrGroup = Get-UsersOrGroupsFromGraph -ObjectId $PrincipalObjectId
        $PrincipalDisplayName = $userOrGroup.DisplayName
        $PrincipalEmail = $userOrGroup.Mail

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/modifyPermissions?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        #Construct the body
        $requestbody = $null

        $requestbody = @{
            put = @(
                @{
                    properties = @{
                        principal = @{
                            email = $PrincipalEmail
                            id = $PrincipalObjectId
                            type = $PrincipalType
                            displayName = $PrincipalDisplayName
                        }
                        roleName = $RoleName
                    }
                }
            )
        }

        $result = InvokeApi -Method POST -Route $route -Body $requestbody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Remove-AdminFlowOwnerRole
{
<#
 .SYNOPSIS
 Removes owner permissions to the flow.
 .DESCRIPTION
 The Remove-AdminFlowOwnerRole sets up permission to flow depending on parameters.
 Use Get-Help Remove-AdminFlowOwnerRole -Examples for more detail.
 .PARAMETER EnvironmentName
 The environment of the flow.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER RoleId
 Specifies the role id of user or group or tenant.
 .EXAMPLE
 Remove-AdminFlowOwnerRole -EnvironmentName "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" -FlowName $flow.FlowName -RoleId "/providers/Microsoft.ProcessSimple/environments/Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877/flows/791fc889-b9cc-4a76-9795-ae45f75d3e48/permissions/1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488"
 deletes flow permision for the given RoleId, FlowName and Environment name.
 #>
    [CmdletBinding(DefaultParameterSetName="Owner")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Owner")]
        [string]$ApiVersion = "2016-11-01",

        [Parameter(Mandatory = $true, ParameterSetName = "Owner", ValueFromPipelineByPropertyName = $true)]
        [string]$FlowName,

        [Parameter(Mandatory = $true, ParameterSetName = "Owner", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Owner", ValueFromPipelineByPropertyName = $true)]
        [string]$RoleId
    )

    process
    {
        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/modifyPermissions?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $requestbody = $null

        $requestbody = @{
            delete = @(
                @{
                    id = $RoleId
                    }
                )
                }

        $result = InvokeApi -Method POST -Route $route -Body $requestbody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

# internal, helper function
function CreateFlowSuspensionReportObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$FlowObj,

        [Parameter(Mandatory = $false)]
        [object]$SuspensionTime,

        [Parameter(Mandatory = $false)]
        [object]$LastRunTime
    )

    if([string]::IsNullOrEmpty($LastRunTime)){
        $LastRunTime = 'No runs in the last 30 days'
    }

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name FlowName -Value $FlowObj.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name Enabled -Value ($FlowObj.properties.state -eq 'Started') `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $FlowObj.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $FlowObj.properties.createdTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedTime -Value $FlowObj.properties.lastModifiedTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $FlowObj.properties.environment.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name SuspensionReason -Value $FlowObj.properties.estimatedSuspensionData.reason `
        | Add-Member -PassThru -MemberType NoteProperty -Name SuspensionTime -Value $SuspensionTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name OwnerHasPowerAppsLicense -Value ($FlowObj.properties.estimatedSuspensionData.powerAppPlanExcluded -eq $true)`
        | Add-Member -PassThru -MemberType NoteProperty -Name OwnerHasDynamicsLicense -Value ($FlowObj.properties.estimatedSuspensionData.dynamicsPlanExcluded -eq $true)`
        | Add-Member -PassThru -MemberType NoteProperty -Name LastRunTime -Value $LastRunTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name Owner -Value $FlowObj.properties.creator.objectId;
}

function CreateAzFlowObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$FlowObj
    )

    $environmentName = $FlowObj.properties.environment.name;
    $flowName = $FlowObj.name;

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $FlowObj.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name FlowName -Value $flowName `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $environmentName;
}

function ShouldShowEnforcement
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$FlowObj,

        [Parameter(Mandatory = $false)]
        [array]$EnforceDates
    )

    if($FlowObj.properties.state -ne 'Started')
    {
      return $false
    }

    $matchedEnforcement = $EnforceDates | Where-Object {`
      $_.Reason.Contains($FlowObj.properties.estimatedSuspensionData.reason) -and $_.OutOfContext -eq (($FlowObj.properties.estimatedSuspensionData.powerAppPlanExcluded -eq $true) -or ($FlowObj.properties.estimatedSuspensionData.dynamicsPlanExcluded -eq $true))`
    } | Select-Object -First 1

    return $null -eq $matchedEnforcement -or $matchedEnforcement.EnabledDate -lt [datetime]::UtcNow
}

function Get-EnforceDates
{
    $enforceDates = @(
      @{
        Reason = @('MissingPremiumLicense', 'MissingAttendedRPA', 'MissingUnAttendedRPA', 'MissingRPAUnattendedAddOn')
        OutOfContext = $false
        Message = 'Premium flows are being used without a Premium license assigned to the flow owner.'
        Recommendation = "For Suspension reason MissingPremiumLicense, assign a Power Automate Premium/Power Automate with attended RPA/Power Automate per user license to the flow owner.`r`n" +
          'For Suspension reason MissingAttendedRPA assign a Power Automate Premium/Power Automate with attended RPA license to the flow owner.`r`n' +
          'For Suspension reason MissingUnattendedRPA, MissingUnattendedAddon, assign a Power Automate Process license to the flow or assign PowerAutomateUnattendedRPA addon to the environment.'
        EnabledDate = [datetime]::Parse('2023-03-03')
      },
      @{
        Reason = @('MissingPremiumLicense', 'MissingAttendedRPA', 'MissingUnAttendedRPA', 'MissingRPAUnattendedAddOn')
        OutOfContext = $true
        Message = "Premium flows created by Power Apps premium and Dynamics 365 licensed users must be in context of the Power App/Dynamics365 application.`r`n" +
          "Power Apps premium and Dynamics 365 licenses provide limited Power Automate usage rights [https://aka.ms/pacontextrights]."
        Recommendation = 'If a premium flow is not triggered by a Power App/not using Dynamics entities and is not linked to any Power App/Dynamics 365 app, you must purchase a  Power Automate Premium license. You can associate a flow to app following: [https://aka.ms/powerautomateassociateapp].'
        EnabledDate = [datetime]::Parse('2023-03-03')
      },
      @{
        Reason = @('NonCompliantServicePrincipalFlow')
        OutOfContext = $false
        Message = 'Premium flows are owned by Application user(SPN) but do not have appropriate license'
        Recommendation = 'For Suspension reason NonCompliantServicePrincipalFlow, assign a Power Automate Process license or per flow license to the flow.'
        EnabledDate = [datetime]::Parse('2023-06-30')
      }
    );

    Write-Host -ForegroundColor Black -BackgroundColor Green 'Non-compliant flow report release dates below. As new enforcements launch, the release dates will be added.'
    return $EnforceDates | ForEach-Object {`
      Write-Host -ForegroundColor Yellow "$($_.EnabledDate): $($_.Message)"
      Write-Host -ForegroundColor White $_.Recommendation
      return $_
    }
}

function Get-AdminFlowAtRiskOfSuspension
{
 <#
 .SYNOPSIS
 Get flows that are at risk of suspension.
 .DESCRIPTION
 The Get-AdminFlowAtRiskOfSuspension gets flows that are at risk of suspension.
 Use Get-Help Get-AdminFlowAtRiskOfSuspension -Examples for more detail.
 .PARAMETER EnvironmentName
 Specifies the environment name.
 .EXAMPLE
 Get-AdminFlowAtRiskOfSuspension -EnvironmentName  7b7ba678-637d-ec84-9bdc-32f4265f449b
 Get flows that are at risk of suspension in environment "7b7ba678-637d-ec84-9bdc-32f4265f449b".
.EXAMPLE
 Get-AdminFlowAtRiskOfSuspension -EnvironmentName  7b7ba678-637d-ec84-9bdc-32f4265f449b -ApiVersion '2016-11-01' | Export-Csv -Path outputDemo.csv -NoTypeInformation
 Export the flows that are at risk of suspension in environment "7b7ba678-637d-ec84-9bdc-32f4265f449b" to file outputDemo.csv.
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $enforceDates = Get-EnforceDates

        Write-Host -ForegroundColor Black -BackgroundColor Green "It can take up to 24 hours for a license change to take effect in the report."

        $routeFlows = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/scheduledSuspensionFlows?api-version={apiVersion}&`$top=50" `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName

        $detectedFlows = (InvokeApi -Method GET -Route $routeFlows -ApiVersion $ApiVersion -ThrowOnFailure -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)).Value `
            | Where-Object { ShouldShowEnforcement -FlowObj $_ -EnforceDates $enforceDates }

        if($detectedFlows.Length -eq 0) {
            Write-Host -ForegroundColor Green "No flows are found at risk of suspension."
        }
        else {
            $detectedFlows | ForEach-Object {`
                $flowResult = $_
                if ($null -ne $flowResult.properties.estimatedSuspensionData) {
                    $routeRuns =  "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/runs?api-version={apiVersion}&`$top=1" `
                    | ReplaceMacro -Macro "{flowName}" -Value $_.name `
                    | ReplaceMacro -Macro "{environment}" -Value $_.properties.environment.name `
                    | ReplaceMacro -Macro "{apiVersion}"  -Value $ApiVersion `
                    | ReplaceMacro -Macro "{flowEndpoint}" -Value $global:currentSession.flowEndpoint;

                    $LastRunTime = $null
                    try {
                        $flowRuns = Invoke-Request `
                            -Uri $routeRuns `
                            -Method GET `
                            -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);

                        $content = ConvertFrom-Json($flowRuns.Content)
                        if ($Content.value.Count -gt 0) {
                            $LastRunTime = $content.value[0].properties.startTime;
                        }
                    }
                    catch {
                        Write-Host "Error while getting flow's last runtime."
                        $LastRunTime = "Failed to retrieve last run time."
                        $response = $_.Exception.Response
                        if ($_.ErrorDetails) {
                            $errorResponse = ConvertFrom-Json $_.ErrorDetails;
                            $code = $response.StatusCode
                            $message = $errorResponse.error.message
                            Write-Verbose "Status Code: '$code'. Message: '$message'"
                        }
                    }

                    $SuspensionTime = $flowResult.properties.estimatedSuspensionData.time
                    if (($null -eq $SuspensionTime) -or ($SuspensionTime -eq "") -or ($flowResult.properties.estimatedSuspensionData.isVisible -ne $true)) {
                      $SuspensionTime = 'N/A'
                    }

                    CreateFlowSuspensionReportObject -FlowObj $flowResult -LastRunTime $LastRunTime -SuspensionTime $SuspensionTime
                }
            }
        }
    }
}

function Remove-AdminFlowApprovals
{
 <#
 .SYNOPSIS
 Removes all active and inactive Flow Approvals.
 .DESCRIPTION
 The Remove-AdminFlowApprovals removes all Approval
 Use Get-Help Remove-AdminFlowApprovals -Examples for more detail.
 .PARAMETER EnvironmentName
 Limits approvals deleted to the specified environment
 .EXAMPLE
 Remove-AdminFlowApprovals -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Finds all approvals assigned to the user in the current environment.
 #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $currentEnvironment = ResolveEnvironment -OverrideId $EnvironmentName;

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environmentName}/users/{currentUserId}/approvals" `
        | ReplaceMacro -Macro "{environmentName}" -Value $currentEnvironment `
        | ReplaceMacro -Macro "{currentUserId}" -Value $global:currentSession.UserId;

        $approvalRequests = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($approvalRequests);
    }
}

function Get-AdminFlowOwnerRole
{
<#
    .SYNOPSIS
    Gets owner permissions to the flow.
    .DESCRIPTION
    The Get-AdminFlowOwnerRole
    Use Get-Help Get-AdminFlowOwnerRole -Examples for more detail.
    .PARAMETER EnvironmentName
    The environment of the flow.
    .PARAMETER FlowName
    Specifies the flow id.
    .PARAMETER Owner
    A objectId of the user you want to filter by.
    .EXAMPLE
    Get-AdminFlowOwnerRole -Owner 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    Returns all flow permissions across all environments for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    .EXAMPLE
    Get-AdminFlowOwnerRole -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -Owner 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    Returns all flow permissions within environment with id 3c2f7648-ad60-4871-91cb-b77d7ef3c239 for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    .EXAMPLE
    Get-AdminFlowOwnerRole -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -Owner 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    Returns all flow permissions for the flow with id 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 for the user with an object id of 53c0a918-ce7c-401e-98f9-1c60b3a723b3
    .EXAMPLE
    Get-AdminFlowOwnerRole -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    Returns all permissions for the flow with id 4d1f7648-ad60-4871-91cb-b77d7ef3c239 in environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = "Environment", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string]$FlowName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [Parameter(Mandatory = $true, ParameterSetName = "User")]
        [Parameter(Mandatory = $true, ParameterSetName = "Environment")]
        [string]$Owner,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [Parameter(Mandatory = $false, ParameterSetName = "Environment")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        if (-not [string]::IsNullOrWhiteSpace($FlowName))
        {
            $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/permissions?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

            $flowRoleAssignmentResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            $pattern = BuildFilterPattern -Filter $Owner

            foreach ($flowRole in $flowRoleAssignmentResult.Value)
            {
                if ($pattern.IsMatch($flowRole.properties.principal.id))
                {
                    CreateFlowRoleAssignmentObject -FlowRoleAssignmentObj $flowRole
                }
            }
        }
        else
        {
            $environments = @();

            if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $environments += @{
                    EnvironmentName = $EnvironmentName
                }
            }
            else {
                $environments = Get-AdminPowerAppEnvironment -ReturnCdsDatabaseType $false
            }

            foreach($environment in $environments)
            {
                $flowResult = Get-AdminFlow -EnvironmentName $environment.EnvironmentName

                foreach ($flow in $flowResult)
                {
                    $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/permissions?api-version={apiVersion}" `
                    | ReplaceMacro -Macro "{flowName}" -Value $flow.FlowName `
                    | ReplaceMacro -Macro "{environment}" -Value $environment.EnvironmentName;


                    $flowRoleAssignmentResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

                    $pattern = BuildFilterPattern -Filter $Owner

                    foreach ($flowRole in $flowRoleAssignmentResult.Value)
                    {
                        if ($pattern.IsMatch($flowRole.properties.principal.id ))
                        {
                            CreateFlowRoleAssignmentObject -FlowRoleAssignmentObj $flowRole
                        }
                    }
                }
            }
        }
    }
}


function Get-AdminFlowUserDetails
{
<#
 .SYNOPSIS
 Returns the Flow user details for the input user Id.
 .DESCRIPTION
 The Get-AdminFlowUserDetails returns the values for ConsentTime, ConsentBusinessAppPlatformTime, IsDisallowedForInternalPlans, ObjectId, Puid, ServiceSettingsSelectionTime, and TenantId.
 Use Get-Help Get-AdminFlowUserDetails -Examples for more detail.
 .PARAMETER UserId
 ID of the user query.
 .EXAMPLE
 Get-AdminFlowUserDetails -UserId 7557f390-5f70-4c93-8bc4-8c2faabd2ca0
 Retrieves the user details associated with the user Id 7557f390-5f70-4c93-8bc4-8c2faabd2ca0
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$UserId = $Global:currentSession.userId,


        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/users/{userId}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{userId}" -Value $userId `
        | ReplaceMacro -Macro "{apiVersion}" -Value $ApiVersion;

        $response = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateFlowUserDetailsObject($response)
    }
}

function Remove-AdminFlowUserDetails
{
<#
 .SYNOPSIS
 Removes the Flow user details for the input user Id. It will throw an error if the input user is an owner of any flows in the tenant.
 .DESCRIPTION
 The Remove-AdminFlowUserDetails deletes the Flow user details assocaited with the input user Id from the Flow database.
 Use Get-Help Remove-AdminFlowUserDetails -Examples for more detail.
 .PARAMETER UserId
 Object Id of the user to delete.
 .EXAMPLE
 Remove-AdminFlowUserDetails -UserId 7557f390-5f70-4c93-8bc4-8c2faabd2ca0
 Removes the details associated with the input user Id 7557f390-5f70-4c93-8bc4-8c2faabd2ca0.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$UserId,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$ApiVersion = "2017-05-01"
    )
    process
    {
        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/users/{userId}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{userId}" -Value $userId `
        | ReplaceMacro -Macro "{apiVersion}" -Value $ApiVersion;

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($response.StatusCode -eq 400)
        {
            Write-Error "All flows for this user must be deleted before user details can be deleted"
        }
        else
        {
            CreateHttpResponse($response)
        }
    }
}

function Add-AdminFlowsToSolution
{
<#
.SYNOPSIS
    Add flows to a solution.
.DESCRIPTION
The Add-AdminFlowsToSolution migrates non-solution flows to a solution.
Use Get-Help Add-AdminFlowsToSolution -Examples for more detail.
 .PARAMETER EnvironmentName
 Limits flow addition/migration to a specific environment.
.PARAMETER SolutionId
Specifies the solution id.
.PARAMETER FlowNames
Specifies the flow names to be added to solution.
.PARAMETER Top
Limits the number of flows that can be added to solution.
.PARAMETER ApiVersion
Specifies the API version.
.EXAMPLE
Add-AdminFlowsToSolution -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
Migrates all the non-solution flows in the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 to a default solution.
.EXAMPLE
Add-AdminFlowsToSolution -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -FlowNames a795f437-3b2b-4f56-ad2c-f3675121fe15,f74970cb-a77b-4762-8f87-b7c105f0ae0a,15ee1047-5beb-4225-9d2f-787f8c5a9f21
Migrates flows in the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 to a default solution.
.EXAMPLE
Add-AdminFlowsToSolution -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -SolutionId 55501614-db8a-4537-8d60-26e80cb80372
Migrates all the non-solution flows in the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 to the solution 55501614-db8a-4537-8d60-26e80cb80372.
.EXAMPLE
Add-AdminFlowsToSolution -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -SolutionId 55501614-db8a-4537-8d60-26e80cb80372 -Top 50
Migrates top 50 non-solution flows to solution
#>

    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$EnvironmentName,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$SolutionId,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string[]]$FlowNames,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [AllowNull()]
        [Nullable[System.Int32]]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Top,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $IsVerbosePresent = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true

        if ($FlowNames -eq $null -or $FlowNames.Count -eq 0)
        {
            $topFlowsToRetrieve = 50

            $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/v2/flows?api-version={apiVersion}&exclude=excludeSolutionCloudFlows&`$top={top}" `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
            | ReplaceMacro -Macro "{top}" -Value $topFlowsToRetrieve;

            $flowResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:$IsVerbosePresent

            $FlowNames = $flowResult.value | ForEach {$_.name}
        }

        $migrationResult = @()

        if($FlowNames -eq $null -or $FlowNames.Count -eq 0)
        {
            if($IsVerbosePresent)
            {
                Write-Host "There are no non-solution flows present in the environment $EnvironmentName."
            }

            return $migrationResult
        }

        if([string]::IsNullOrWhiteSpace($SolutionId))
        {
            $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/migrateFlows?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
        }
        else{
            $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/migrateFlows?api-version={apiVersion}&solutionId={solutionId}" `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
            | ReplaceMacro -Macro "{solutionId}" -Value $SolutionId;
        }

        if($Top -ne $null)
        {
            $totalFlowsToMigrate = [math]::Min($Top, $FlowNames.Count)
            $FlowNames = $FlowNames[0..($totalFlowsToMigrate - 1)]
        }
        else
        {
            $totalFlowsToMigrate = $FlowNames.Count
        }

        $batchSize = [math]::Min(20, $totalFlowsToMigrate)

        for($start = 0; $start -lt $totalFlowsToMigrate; $start += $batchSize)
        {
            $end = [math]::Min($start + $batchSize, $totalFlowsToMigrate) - 1
            $flows = $FlowNames[$start..$end]

            if($IsVerbosePresent)
            {
                Write-Host "Total flows to migrate: $totalFlowsToMigrate, Current batch of flows for migration $flows, start index: $start, end index: $end,"
            }

            $requestBody = @{ "flowsToMigrate" = $flows }
            $currentBatchMigrationResult = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Body $requestBody -ThrowOnFailure -Verbose:$IsVerbosePresent

            if($IsVerbosePresent)
            {
               Write-Host "Migration result: $currentBatchMigrationResult"
            }

            $migrationResult += $currentBatchMigrationResult
        }

        return $migrationResult

    }
}

function Start-EUDBMigration
{
<#
.SYNOPSIS
    Initiate migration of flows for the EUDB compliance.
.DESCRIPTION
The Start-EUDBMigration initiates migration of flows present in the user resource group to an environment resource group.
Use Get-Help Start-EUDBMigration -Examples for more detail.
 .PARAMETER EnvironmentName
 Limits migration of flows to a specific environment.
.PARAMETER FlowNames
Specifies the flow names to be migrated
.PARAMETER ApiVersion
Specifies the API version.
.EXAMPLE
Start-EUDBMigration -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
Starts migration of all the non-compliant flows in the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239.
.EXAMPLE
Start-EUDBMigration -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -FlowNames a795f437-3b2b-4f56-ad2c-f3675121fe15,f74970cb-a77b-4762-8f87-b7c105f0ae0a,15ee1047-5beb-4225-9d2f-787f8c5a9f21
Starts migration of flows in the environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239.
#>

    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$EnvironmentName,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string[]]$FlowNames,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$ApiVersion = "2016-11-01"
    )

    process
    {
        $IsVerbosePresent = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/migrateResourceGroup?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `

        $migrationResult = @()

        if ($FlowNames -eq $null -or $FlowNames.Count -eq 0)
        {
            $migrationResult = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:$IsVerbosePresent

            if($IsVerbosePresent)
			{
				Write-Host "Migration result: $migrationResult"
			}
        }
        else
        {
            $totalFlowsToMigrate = $FlowNames.Count
            $batchSize = [math]::Min(20, $totalFlowsToMigrate)

            for($start = 0; $start -lt $totalFlowsToMigrate; $start += $batchSize)

            {
                $end = [math]::Min($start + $batchSize, $totalFlowsToMigrate) - 1
                $flows = $FlowNames[$start..$end]
                if($IsVerbosePresent)
                {
                    Write-Host "Total flows to migrate: $totalFlowsToMigrate, Current batch of flows for migration $flows, start index: $start, end index: $end,"
                }
            
                $requestBody = @{ "flowsToMigrate" = $flows }
                $currentBatchMigrationResult = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Body $requestBody -ThrowOnFailure -Verbose:$IsVerbosePresent
            
                if($IsVerbosePresent)
                {
                    Write-Host "Migration result: $currentBatchMigrationResult"
                }
                $migrationResult += $currentBatchMigrationResult
            }
        }

        return $migrationResult

    }
}

function Cancel-EUDBMigration
{
<#
.SYNOPSIS
    Cancel resource group migration operations for an environment as per EUDB compliance. 
.DESCRIPTION
The Cancel-EUDBMigration cancels resource migration operations for an environment. These operations migrate flows from a user 
resource group to an environment resource group. Cancellation in this context applies to ongoing operations and/or new migration 
operations. Therefore, after configuring an environment for cancellation, Start-EUDBMigration should be used so that new migration 
operations in the environment can be performed.
Use Get-Help Cancel-EUDBMigration -Examples for more details.
.PARAMETER EnvironmentName
The environment guid.
.PARAMETER ApiVersion
Specifies the API version.
.EXAMPLE
Cancel-EUDBMigration -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
Cancels the resource group migration operations, for environment: 3c2f7648-ad60-4871-91cb-b77d7ef3c239.
#>

    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$EnvironmentName,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $IsVerbosePresent = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true

        # Call the cancelRgMigration API
        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environmentName}/cancelRgMigration?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        # Api call
        $cancelRgMigrationResult = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:$IsVerbosePresent

        if($IsVerbosePresent)
		{
			Write-Host "Cancellation result: $cancelRgMigrationResult"
		}

        return $cancelRgMigrationResult
    }
}

function Get-EUDBMigrationStatus
{
<#
.SYNOPSIS
    Checks the status of migration of flows for the EUDB compliance.
.DESCRIPTION
The Get-EUDBMigrationStatus checks the status of migration where flows present in the user resource group should be moved to an environment resource group.
Use Get-Help Get-EUDBMigrationStatus -Examples for more detail.
 .PARAMETER EnvironmentName
 Limits migration of flows to a specific environment.
.PARAMETER OperationId
Specifies the operation Id returned in the response header of the Start-EUDBMigration cmdlet
.PARAMETER ApiVersion
Specifies the API version.
.EXAMPLE
Get-EUDBMigrationStatus -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -OperationId f47b73a4-9cb2-4d6c-a1e2-df8fd93a81be
Checks the migration status of the specified operation Id.
#>

    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$EnvironmentName,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$OperationId,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flow")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $IsVerbosePresent = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/operationTypes/ResourceGroupMigration/v1/operations/{operationId}?api-version={apiVersion}" `
		| ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{operationId}" -Value $OperationId;

		$migrationStatus = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:$IsVerbosePresent

        if($IsVerbosePresent)
        {
            Write-Host "Migration status: $migrationStatus"
        }

		return $migrationStatus
    }
}

function Add-AdminFlowPowerAppContext
{
<#
 .SYNOPSIS
 Add Powerapp context to the specific flow.
 .DESCRIPTION
 The Add-AdminFlowPowerAppContext add Powerapp context to the specific flow.
 Use Get-Help Add-AdminFlowPowerAppContext -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER AppName
 Specifies the PowerApp id.
 .EXAMPLE
 Add-AdminFlowPowerAppContext -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -PowerAppName c3dba9c8-0f42-4c88-8110-04b582f20735
 Add flow 4d1f7648-ad60-4871-91cb-b77d7ef3c239 flow in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" to the context of PowerApp 'c3dba9c8-0f42-4c88-8110-04b582f20735' in the same environment.
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $app = Get-AdminPowerApp -EnvironmentName $EnvironmentName -AppName $AppName

        if($app.AppName -eq $null -or $app.AppName -eq "")
        {
            $route = "https://{powerAppsEndpoint}/environments/{environmentName}/apps/{appName}?api-version={apiVersion}" `
                | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
                | ReplaceMacro -Macro "{appName}" -Value $AppName;

            $appResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            if($appResult.name -eq $null -or $appResult.name -eq "")
            {
                Write-Error "Specified PowerApp '$AppName' does not exist, please specify a valid PowerApp."
                return
            }
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/associateBillingContext?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $body = @{
            contextLinkType = "unifiedApp"
            contextLinkId = $AppName
        }

        $result = InvokeApi -Method POST -Route $route -Body $body -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Get-AdminNonZoneRedundantFlows
{
<#
 .SYNOPSIS
 Lists flows that are not in Azure zone redundant regions.
 .DESCRIPTION
 The Get-AdminNonZoneRedundantFlows returns a list of flow ids that are not in Azure zone redundant regions.
 Use Get-Help Get-AdminNonZoneRedundantFlows -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Get-AdminNonZoneRedundantFlows -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 Returns a array of flow ids for environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877".
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        if($EnvironmentName -eq $null -or $EnvironmentName -eq "")
        {
            Write-Error "Specified environment '$EnvironmentName' does not exist, please specify a valid environment."
            return
        }    
    
        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/nonZoneRedundantFlows?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $result = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($result.value.length -gt 0){
            foreach ($flow in $result.value) {
                if (!($flow.name.EndsWith("-azshadow"))) {
                    CreateAzFlowObject -FlowObj $flow
                }
            }   
        } else {
            Write-Host "No flows found."
        }
        
    }
}

function Get-AdminEnableZoneRedundancyPreflightCheck
{
<#
 .SYNOPSIS
 Checks if the target flow can be migrated without causing disruptions.
 .DESCRIPTION
 The Get-AdminEnableZoneRedundancyPreflightCheck returns a pass or fail status depending on if the targeted flow can be migrated.
 Use Get-Help Get-AdminEnableZoneRedundancyPreflightCheck -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .PARAMETER FlowName
 Specifies the flow id.
 .EXAMPLE
 Get-AdminEnableZoneRedundancyPreflightCheck -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239
 Returns a pass message for flow "4d1f7648-ad60-4871-91cb-b77d7ef3c239" in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877".
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        if($FlowName -eq $null -or $FlowName -eq "")
        {
            Write-Error "Specified flow '$FlowName' does not exist, please specify a valid flow id."
            return
        }    

        if($EnvironmentName -eq $null -or $EnvironmentName -eq "")
        {
            Write-Error "Specified environment '$EnvironmentName' does not exist, please specify a valid environment."
            return
        }    
    
        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/enableZoneRedundancy?api-version={apiVersion}&validation=true"`
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName;

        
        $result = InvokeApiNoParseContent -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        if ($result.StatusCode -eq 200) {
            Write-Host "Pass. Flow $FlowName meets the requirements to be migrated." -ForegroundColor Green
        } else {
            Write-Error "Fail. Flow $FlowName does not meet the requirements to be migrated."
        }

        CreatePreflightResponse($result)
    }
}

function Start-AdminFlowEnableZoneRedundancy 
{
<#
 .SYNOPSIS
 Triggers the Azure zone redundancy migration for the target flow.
 .DESCRIPTION
 The Start-AdminFlowEnableZoneRedundancy  starts migration for the target flow and depending on user input, will return a success message upon migration completion.
 Use Get-Help Start-AdminFlowEnableZoneRedundancy  -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER WaitUntilFinish
 Flag to set whether to wait for the migration operation to complete.
 .EXAMPLE
 Start-AdminFlowEnableZoneRedundancy  -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -WaitUntilFinish True
 Returns a success message for migrating flow "4d1f7648-ad60-4871-91cb-b77d7ef3c239" from environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877".
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [Switch] $WaitUntilFinish,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        if ($FlowName -eq $null -or $FlowName -eq "")
        {
            Write-Error "Specified flow '$FlowName' does not exist, please specify a valid flow id."
            return
        }    

        if ($EnvironmentName -eq $null -or $EnvironmentName -eq "")
        {
            Write-Error "Specified environment '$EnvironmentName' does not exist, please specify a valid environment."
            return
        }    
    
        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/enableZoneRedundancy?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName;

        $result = InvokeApiNoParseContent -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($WaitUntilFinish)
        {
            while($result.StatusCode -eq 202)
            {
                Start-Sleep -s 5
                $migrationStatusUri = $result.Headers['Location'] | Select-Object -First 1
                $result = InvokeApiNoParseContent -Method GET -Route $migrationStatusUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            }
        }

        $statusMessage = Get-StatusMessage -StatusCode $result.StatusCode
        Write-Host $statusMessage

        $operation = CreateMigrationOperation -ResponseObject $result
        $operation | Out-Default
        return $operation
    }
}

function Get-AdminEnableZoneRedundancyStatus
{
<#
 .SYNOPSIS
 Checks the status of the flow Azure zone redundant region migration.
 .DESCRIPTION
 The Get-AdminEnableZoneRedundancyStatus returns success or failure depending on if the migration is complete.
 Use Get-Help Get-AdminEnableZoneRedundancyStatus -Examples for more detail.
 .PARAMETER MigrationOperation
 Specifies the status code and next link associated with the migration operation.
 .EXAMPLE
 $operation = New-AdminFlowEnableZoneRedundancy  -EnvironmentName e49fd0a4-47eb-e1f3-b89e-0b8d7d6f3471 -FlowName d354db06-448c-45d7-8f22-14569c4913cd -WaitUntilFinish false
 Get-AdminEnableZoneRedundancyStatus -MigrationOperation $operation
 Returns success
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [object] $MigrationOperation
    )
    process
    {
        if ($MigrationOperation -eq $null)
        {
            Write-Error "Specified migration operation does not exist, please input a valid migration operation."
            return
        }   

        if (!$MigrationOperation.NextLink)
        {
            $statusMessage = Get-StatusMessage -StatusCode $MigrationOperation.Status
            Write-Host $statusMessage
            return
        }  

        $result = InvokeApiNoParseContent -Method GET -Route $MigrationOperation.nextLink -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        $statusMessage = Get-StatusMessage -StatusCode $result.StatusCode
        Write-Host $statusMessage

        CreateMigrationOperation -ResponseObject $result
    }
}

function Get-AdminFlowEnvironmentCmkStatus
{
<#
 .SYNOPSIS
 Get CMK Status for flow environment.
 .DESCRIPTION
 The Get-AdminFlowEnvironmentCmkStatus get CMK status for flow environment.
 Use Get-Help Get-AdminFlowEnvironmentCmkStatus -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Get-AdminFlowEnvironmentCmkStatus -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 Gets the CMK status for environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877".
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/environments/{environment}/cmkstatus?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $result = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Get-AdminFlowEncryptedByMicrosoftKey
{
<#
 .SYNOPSIS
 Lists flows that are encrypted by microsoft key.
 .DESCRIPTION
 The Get-AdminFlowEncryptedByMicrosoftKey returns a list of flow ids that are not in CMK migrated.
 Use Get-Help Get-AdminFlowEncryptedByMicrosoftKey -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Get-AdminFlowEncryptedByMicrosoftKey -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 Returns a array of flow ids for environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" that is still encrypted by Microsoft Managed Key.
 .EXAMPLE
 Get-AdminFlowEncryptedByMicrosoftKey -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -ListCannotMigrateToCustomerManagedKey
 Only returns a array of flow ids for environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" that cannot be migrated to be encrypted by Customer Managed Key.
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [Switch]$ListCannotMigrateToCustomerManagedKey,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        if($EnvironmentName -eq $null -or $EnvironmentName -eq "")
        {
            Write-Error "Specified environment '$EnvironmentName' does not exist, please specify a valid environment."
            return
        }

        if ($ListCannotMigrateToCustomerManagedKey)
        {
            $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/nonCmkMigratedFlows?api-version={apiVersion}&exclude=cmkMigrationSupportedFlows" `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;
        }
        else
        {
            $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/nonCmkMigratedFlows?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;
        }

        $result = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($result.value.length -gt 0){
            foreach ($flow in $result.value) {
                CreateAzFlowObject -FlowObj $flow
            }
        } else {
            Write-Host "No flows found."
        }
    }
}

function Get-AdminFlowWithHttpAction
{
<#
 .SYNOPSIS
 Lists flows with http action.
 .DESCRIPTION
 The Get-AdminFlowWithHttpAction returns a list of flow ids with http action in the flow definition.
 Use Get-Help Get-AdminFlowWithHttpAction -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Get-AdminFlowWithHttpAction -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 Returns a array of flow ids for environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" that contains http action in the flow definition.
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        if($EnvironmentName -eq $null -or $EnvironmentName -eq "")
        {
            Write-Error "Specified environment '$EnvironmentName' does not exist, please specify a valid environment."
            return
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/httpActionFlows?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $result = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($result.value.length -gt 0){
            foreach ($flow in $result.value) {
                CreateAzFlowObject -FlowObj $flow
            }
        } else {
            Write-Host "No flows found."
        }
    }
}

function Get-AdminFlowWithMigratingTriggerUrl
{
<#
 .SYNOPSIS
 Lists flows with migrating trigger URL.
 .DESCRIPTION
 The Get-AdminFlowWithMigratingTriggerUrl returns a list of flow ids that require an update to the trigger url.
 Use Get-Help Get-AdminFlowWithMigratingTriggerUrl -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .EXAMPLE
 Get-AdminFlowWithMigratingTriggerUrl -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877
 Returns a array of flow ids for environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" that contains flows with migrating triggers in the flow definition.
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        if($EnvironmentName -eq $null -or $EnvironmentName -eq "")
        {
            Write-Error "Specified environment '$EnvironmentName' does not exist, please specify a valid environment."
            return
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flowsWithMigratingTriggerUrl?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $result = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($result.value.length -gt 0){
            foreach ($flow in $result.value) {
                CreateAzFlowObject -FlowObj $flow
            }
        } else {
            Write-Host "No flows found."
        }
    }
}

function Remove-AdminFlowPowerAppContext
{
<#
 .SYNOPSIS
 Remove Powerapp context to the specific flow.
 .DESCRIPTION
 The Remove-AdminFlowPowerAppContext remove Powerapp context from the specific flow.
 Use Get-Help Remove-AdminFlowPowerAppContext -Examples for more detail.
 .PARAMETER EnvironmentName
 Limit apps returned to those in a specified environment.
 .PARAMETER FlowName
 Specifies the flow id.
 .PARAMETER AppName
 Specifies the PowerApp id.
 .EXAMPLE
 Remove-AdminFlowPowerAppContext -EnvironmentName Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877 -FlowName 4d1f7648-ad60-4871-91cb-b77d7ef3c239 -PowerAppName c3dba9c8-0f42-4c88-8110-04b582f20735
 Remove flow 4d1f7648-ad60-4871-91cb-b77d7ef3c239 flow in environment "Default-55abc7e5-2812-4d73-9d2f-8d9017f8c877" from the context of PowerApp 'c3dba9c8-0f42-4c88-8110-04b582f20735' in the same environment.
 #>
    [CmdletBinding(DefaultParameterSetName="Flow")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $FlowName,

        [Parameter(Mandatory = $true, ParameterSetName = "Flow", ValueFromPipelineByPropertyName = $true)]
        [string] $AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "Flow")]
        [string] $ApiVersion = "2016-11-01"
    )
    process
    {
        if ($ApiVersion -eq $null -or $ApiVersion -eq "")
        {
            Write-Error "Api Version must be set."
            throw
        }

        $route = "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/{environment}/flows/{flowName}/dissociateBillingContext?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{flowName}" -Value $FlowName `
        | ReplaceMacro -Macro "{environment}" -Value (ResolveEnvironment -OverrideId $EnvironmentName);

        $body = @{
            contextLinkType = "unifiedApp"
            contextLinkId = $AppName
        }

        $result = InvokeApi -Method POST -Route $route -Body $body -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($result)
    }
}

function Set-AdminPowerAppAsFeatured
{
<#
 .SYNOPSIS
 Updates the input PowerApp to be a featured application for the tenant.
 .DESCRIPTION
 The Set-AdminPowerAppAsFeatured changes the isFeaturedApp flag of a PowerApp to true.
 Use Get-Help Set-AdminPowerAppAsFeatured -Examples for more detail.
 .PARAMETER AppName
 App Id of PowerApp to operate on.
 .PARAMETER EnvironmentName
 The Environment's identifier.
 .PARAMETER ApiVersion
 PowerApps Api version date, defaults to "2017-05-01"
 .PARAMETER ForceLease
 Forces the lease when overwriting the PowerApp fields. Defaults to false if no input is provided.
 .EXAMPLE
 Set-AdminPowerAppAsFeatured -PowerAppName c3dba9c8-0f42-4c88-8110-04b582f20735
 Updates the input PowerApp to be a featured application of that tenant.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$AppName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName = $null,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2017-05-01",

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Boolean]$ForceLease
    )
    process
    {
        if(-not [string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $requestBody = @{
                "isFeaturedApp" = $true
            }

            $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setFeaturedApp?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
            | ReplaceMacro -Macro "{appName}" -Value $AppName;
            InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }
        else
        {
            $getPowerAppUri = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/apps/{appName}`?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{appName}" -Value $AppName;

            $powerApp = InvokeApi -Route $getPowerAppUri -Method Get -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            $powerApp.properties.isFeaturedApp = $true

            AcquireLeaseAndPutApp -AppName $AppName -ApiVersion $ApiVersion -PowerApp $powerApp -ForceLease $ForceLease
        }
    }
}

function Clear-AdminPowerAppAsFeatured
{
<#
 .SYNOPSIS
 Removes the input PowerApp as a featured application for the tenant. The input app must not be set as a hero app to unset it as a featured app.
 .DESCRIPTION
 The Clear-AdminPowerAppAsFeatured changes the isFeaturedApp flag of a PowerApp to false.
 Use Get-Help Clear-AdminPowerAppAsFeatured -Examples for more detail.
 .PARAMETER AppName
 App Id of PowerApp to operate on.
 .PARAMETER EnvironmentName
 The Environment's identifier.
 .PARAMETER ApiVersion
 PowerApps Api version date, defaults to "2017-05-01"
 .PARAMETER ForceLease
 Forces the lease when overwriting the PowerApp fields. Defaults to false if no input is provided.
 .EXAMPLE
 Clear-AdminPowerAppAsFeatured -PowerAppName c3dba9c8-0f42-4c88-8110-04b582f20735
 Updates the input PowerApp to be a regular (not featured) application of that tenant.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$AppName,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName = $null,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2017-05-01",

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Boolean]$ForceLease
    )
    process
    {
        if (-not [string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $requestBody = @{
                "isFeaturedApp" = $false
            }

            $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setFeaturedApp?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
            | ReplaceMacro -Macro "{appName}" -Value $AppName;
            InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }
        else
        {
            $getPowerAppUri = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/apps/{appName}`?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{appName}" -Value $AppName;

            $powerApp = InvokeApi -Route $getPowerAppUri -Method Get -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            if($powerApp.properties.isHeroApp -eq $true)
            {
                Write-Error "Must unset the app as a Hero before unsetting the app as featured."
                return
            }

            $powerApp.properties.isFeaturedApp = $false

            AcquireLeaseAndPutApp -AppName $AppName -ApiVersion $ApiVersion -PowerApp $powerApp -ForceLease $ForceLease
        }
    }
}

function Set-AdminPowerAppAsHero
{
<#
 .SYNOPSIS
 Identifies the input PowerApp as a hero application. The input app must be set as a featured app to be set as a hero.
 .DESCRIPTION
 The Set-AdminPowerAppAsHero changes the isHero flag of a PowerApp to true.
 Use Get-Help Set-AdminPowerAppAsHero -Examples for more detail.
 .PARAMETER AppName
 App Id of PowerApp to operate on.
 .PARAMETER ApiVersion
 PowerApps Api version date, defaults to "2017-05-01"
 .PARAMETER ForceLease
 Forces the lease when overwriting the PowerApp fields. Defaults to false if no input is provided.
 .EXAMPLE
 Set-AdminPowerAppAsHero -PowerAppName c3dba9c8-0f42-4c88-8110-04b582f20735
 Updates the input PowerApp to be the hero application of that tenant.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2017-05-01",

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Boolean]$ForceLease
    )
    process
    {
        $getPowerAppUri = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/apps/{appName}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        $powerApp = InvokeApi -Route $getPowerAppUri -Method Get -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if($powerApp.properties.isFeaturedApp -ne $true)
        {
            Write-Error "Must set the app as a Featured app before setting it as a Hero."
            return
        }

        $powerApp.properties.isHeroApp = $true

        AcquireLeaseAndPutApp -AppName $AppName -ApiVersion $ApiVersion -PowerApp $powerApp -ForceLease $ForceLease
    }
}

function Clear-AdminPowerAppAsHero
{
<#
 .SYNOPSIS
 Removes the input PowerApp as a hero application.
 .DESCRIPTION
 The Clear-AdminPowerAppAsHero changes the isHero flag of a PowerApp to false.
 Use Get-Help Clear-AdminPowerAppAsHero -Examples for more detail.
 .PARAMETER AppName
 App Id of PowerApp to operate on.
 .PARAMETER ApiVersion
 PowerApps Api version date, defaults to "2017-05-01"
 .PARAMETER ForceLease
 Forces the lease when overwriting the PowerApp fields. Defaults to false if no input is provided.
 .EXAMPLE
 Clear-AdminPowerAppAsHero -PowerAppName c3dba9c8-0f42-4c88-8110-04b582f20735
 Updates the input PowerApp to be a regular application.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Name")]
        [string]$AppName,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2017-05-01",

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Boolean]$ForceLease
    )
    process
    {
        $getPowerAppUri = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/apps/{appName}`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        $powerApp = InvokeApi -Route $getPowerAppUri -Method Get -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        $powerApp.properties.isHeroApp = $false

        AcquireLeaseAndPutApp -AppName $AppName -ApiVersion $ApiVersion -PowerApp $powerApp -ForceLease $ForceLease
    }
}


function Set-AdminPowerAppApisToBypassConsent
{
<#
 .SYNOPSIS
 Sets the consent bypass flag so users are not required to authorize API connections for the input PowerApp.
 .DESCRIPTION
 The Set-AdminPowerAppApisToBypassConsent changes the bypassConsent flag of a Power App to true. If the Power App is a model driven app, please use the name of the app instead of the app id.
 Use Get-Help Set-AdminPowerAppApisToBypassConsent -Examples for more detail.
 .PARAMETER EnvironmentName
 Environment id of the PowerApp.
 .PARAMETER AppName
 App Id of PowerApp to operate on.
 .PARAMETER ApiVersion
 PowerApps Api version date, defaults to "2021-02-01"
 .EXAMPLE
 Set-AdminPowerAppApisToBypassConsent -EnvironmentName 08b4e32a-4e0d-4a69-97da-e1640f0eb7b9 -AppName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Updates the input PowerApp to not require consent for APIs in the production tenant of the logged in user.
 .EXAMPLE
 Set-AdminPowerAppApisToBypassConsent -EnvironmentName 08b4e32a-4e0d-4a69-97da-e1640f0eb7b9 -AppName msdynce_saleshub
 Updates the msdynce_saleshub model driven app to not require consent for APIs in the production tenant of the logged in user.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
            "bypassConsent" = $true
        }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setPowerAppConnectionDirectConsentBypass?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        return InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Clear-AdminPowerAppApisToBypassConsent
{
<#
 .SYNOPSIS
 Removes the consent bypass so users are required to authorize API connections for the input PowerApp.
 .DESCRIPTION
 The Clear-AdminPowerAppApisToBypassConsent changes the bypassConsent flag of a PowerApp to false.
 Use Get-Help Clear-AdminPowerAppApisToBypassConsent -Examples for more detail.
 .PARAMETER AppName
 App Id of PowerApp to operate on.
 .PARAMETER EnvironmentName
 Environment id of the PowerApp.
 .PARAMETER ApiVersion
 PowerApps Api version date, defaults to "2021-02-01"
 .EXAMPLE
 Clear-AdminPowerAppApisToBypassConsent -EnvironmentName 08b4e32a-4e0d-4a69-97da-e1640f0eb7b9 -AppName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Updates the input PowerApp to require consent in the production tenant of the logged in user.
 #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
            "bypassConsent" = $false
        }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setPowerAppConnectionDirectConsentBypass?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        return InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-AdminPowerAppConditionalAccessAuthenticationContextIds
{
    <#
    .SYNOPSIS
    Gets the authentication context ids attached to the Power App.
    .DESCRIPTION
    The Get-AdminPowerAppConditionalAccessAuthenticationContextIds information about all authentication context ids attached to an input Power App.
    These values are used to check against conditional access policies set by the Admin in the tenant.
    Use Get-Help Get-AdminPowerAppConditionalAccessAuthenticationContextIds -Examples for more detail.
    .PARAMETER EnvironmentName
    Environment where the input Power App is located.
    .PARAMETER AppName
    PowerApp to list connectors for.
    .EXAMPLE
    Get-AdminPowerAppConditionalAccessAuthenticationContextIds -EnvironmentName 643268a6-c680-446f-b8bc-a3ebbf98895f -AppName fc947231-728a-4a74-a654-64b0f22a0d71
    Returns all authentication context ids attached to the Power App 'fc947231-728a-4a74-a654-64b0f22a0d71' if any.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/getPowerAppConditionalAccessDetails?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        return InvokeApi -Route $route -Method GET -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-AdminPowerAppConditionalAccessAuthenticationContextIds
{
    <#
    .SYNOPSIS
    Sets the authentication context ids for the input Power App.
    .DESCRIPTION
    The Set-AdminPowerAppConditionalAccessAuthenticationContextIds sets the authentication context ids for the input Power App.
    These values are used to check against conditional access policies set by the Admin in the tenant.
    Use Get-Help Set-AdminPowerAppApisToBypassConsent -Examples for more detail.
    .PARAMETER EnvironmentName
    Environment id of the Power App.
    .PARAMETER AppName
    App Id of PowerApp to operate on.
    .PARAMETER AuthenticationContextIds
    Authentication context ids to be associated with the Power App.
    .PARAMETER ApiVersion
    Power Apps Api version date, defaults to "2021-02-01"
    .EXAMPLE
    Set-AdminPowerAppConditionalAccessAuthenticationContextIds -EnvironmentName 08b4e32a-4e0d-4a69-97da-e1640f0eb7b9 -AppName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -AuthenticationContextIds c1,c2
    Sets the authentication context ids to array with values 'c1,c2' for the input Power App '3c2f7648-ad60-4871-91cb-b77d7ef3c239'
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$AuthenticationContextIds,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
            "authContextIds" = $AuthenticationContextIds
        }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setPowerAppConditionalAccessDetails?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        return InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-AdminPowerAppConditionalAccessAuthenticationContextIds
{
    <#
    .SYNOPSIS
    Removes the Authentication context ids associated with the input PowerApp.
    .DESCRIPTION
    The Remove-AdminPowerAppConditionalAccessAuthenticationContextIds removes the Authentication context ids from the  input Power App metadata.
    These values are used to check against conditional access policies set by the Admin in the tenant.
    Use Get-Help Remove-AdminPowerAppConditionalAccessAuthenticationContextIds -Examples for more detail.
    .PARAMETER EnvironmentName
    Environment id of the Power App.
    .PARAMETER AppName
    App Id of PowerApp to operate on.
    .PARAMETER ApiVersion
    PowerApps Api version date, defaults to "2021-02-01"
    .EXAMPLE
    Remove-AdminPowerAppConditionalAccessAuthenticationContextIds -EnvironmentName 08b4e32a-4e0d-4a69-97da-e1640f0eb7b9 -AppName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
    Removes the authentication context ids associated with the input PowerApp '3c2f7648-ad60-4871-91cb-b77d7ef3c239'.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
            "authContextIds" = @()
        }

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentName}/apps/{appName}/setPowerAppConditionalAccessDetails?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

        return InvokeApi -Route $route -Body $RequestBody -Method POST -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-AdminVirtualConnectors
{
    <#
    .SYNOPSIS
    Retrieves virtual connectors available to the tenant.
    .DESCRIPTION
    Get-AdminVirtualConnectors cmdlet gets virtual connectors for the tenant.
    Use Get-Help Get-AdminVirtualConnectors -Examples for more detail.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-AdminVirtualConnectors
    Retrieves virtual connectors in the tenant.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
		Test-PowerAppsAccount;

        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/connectors/metadata/virtual" `
        | ReplaceMacro -Macro "{bapEndpoint}" -Value $global:currentSession.bapEndpoint;

        return Invoke-Request -Uri $route -Method GET -Body $null -ParseContent -ThrowOnFailure
    }
}

function Get-AdminDlpPolicy
{
    <#
    .SYNOPSIS
    Retrieves api policy objects and provides the option to print out the connectors in each data group.
    .DESCRIPTION
    Get-AdminDlpPolicy cmdlet gets policy objects for the logged in admin's tenant.
    Use Get-Help Get-AdminDlpPolicy -Examples for more detail.
    .PARAMETER PolicyName
    Retrieves the policy with the input name (identifier).
    .PARAMETER ShowHbi
    Prints out the hbi/business data group api connections if true.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-AdminDlpPolicy
    Retrieves all policies in the tenant.
    .EXAMPLE
    Get-AdminDlpPolicy -PolicyName 78d6c98c-aaa0-4b2b-91c3-83d211754d8a
    Retrieves details on the policy 78d6c98c-aaa0-4b2b-91c3-83d211754d8a.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Filter", ValueFromPipelineByPropertyName = $true)]
        [string]$Filter,

        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "Filter", ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $false)]
        [string]$CreatedBy,

        [Parameter(Mandatory = $false, ParameterSetName = "Filter")]
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [object]$ApiVersion = "2016-11-01"
    )
    process
    {
        # get all policies
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{apiVersion}" -Value $ApiVersion;

        $response = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        # filter and returns policies that match parameters
        Get-FilteredApiPolicies -PolicyName $PolicyName -ApiPolicyResult $response -Filter $Filter -CreatedBy $CreatedBy
    }
}

function New-AdminDlpPolicy
{
    <#
    .SYNOPSIS
    Creates and inserts a new api policy into the tenant. By default the environment filter is off, and all api connections are in the no business data group (lbi).
    .DESCRIPTION
    New-AdminDlpPolicy cmdlet creates a new DLP policy for the logged in admin's tenant.
    Use Get-Help New-AdminDlpPolicy -Examples for more detail.
    .PARAMETER DisplayName
    Creates the policy with the input display name.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER BlockNonBusinessDataGroup
    Block non business data group.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .PARAMETER SchemaVersion
    Specifies the schema version to use, 2016-11-01 or 2018-11-01 (HTTP connectors included).
    .EXAMPLE
    New-AdminDlpPolicy -DisplayName "MetroBank Policy"
    Creates a new policy with the display name 'MetroBank Policy' in the tenant.
    .EXAMPLE
    New-AdminDlpPolicy -DisplayName "MetroBank Policy" -EnvironmentName Default-02c201b0-db76-4a6a-b3e1-a69202b479e6
    Creates a new policy with the display name 'MetroBank Policy' in the environment Default-02c201b0-db76-4a6a-b3e1-a69202b479e6.
    .EXAMPLE
    New-AdminDlpPolicy -DisplayName "MetroBank Policy" -BlockNonBusinessDataGroup $true
    Creates a new policy with the display name 'MetroBank Policy' in the tenant and blocks lbi data group.
    .EXAMPLE
    New-AdminDlpPolicy -DisplayName "MRA Digital" -SchemaVersion 2018-11-01
    Creates a new policy with the display name 'MRA Digital' and schema version '2018-11-01' (includes HTTP connectors).
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$DisplayName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [bool]$BlockNonBusinessDataGroup = $false,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01",

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [ValidateSet("2016-10-01-preview", "2018-11-01")][string]$SchemaVersion = "2016-10-01-preview"
    )
    process
    {
        if([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $createApiPolicyRoute = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies?api-version={apiVersion}"
            $type = "Microsoft.BusinessAppPlatform/scopes/apiPolicies"
        }
        else
        {
            $createApiPolicyRoute = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/apiPolicies?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
            $type = "Microsoft.BusinessAppPlatform/scopes/environments/apiPolicies"
        }

        $schema = "https://schema.management.azure.com/providers/Microsoft.BusinessAppPlatform/schemas/{schemaVersion}/apiPolicyDefinition.json#" `
        | ReplaceMacro -Macro "{schemaVersion}" -Value $SchemaVersion;

        $rules = @{
                    dataFlowRule = @{
                        actions = @{
                            blockAction = @{
                                type = "Block"
                            }
                        }
                        parameters = @{
                            destinationApiGroup = "lbi"
                            sourceApiGroup = "hbi"
                        }
                        type = "DataFlowRestriction"
                    }
                }

        $constraints = @{}
        if(-not [string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $getEnvironmentUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}?`$expand=permissions&api-version={apiVersion}" `
                | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName `
                | ReplaceMacro -Macro "{apiVersion}" -Value $ApiVersion

            $environmentInput = InvokeApi -Method GET -Route $getEnvironmentUri -ApiVersion $ApiVersion  -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)`
                | %{ New-Object -TypeName PSObject -Property @{ id = $_.id; name = $_.name; type = $_.type } }

            $constraints = @{
                environmentFilter1 = @{
                    parameters = @{
                        environments = @($environmentInput)
                        filterType = "include"
                    }
                    type = "environmentFilter"
                }
            }
        }

        $hbiApis = @()
        $lbiDescription = "No business data allowed"
        if ($BlockNonBusinessDataGroup -eq $true)
        {
            if([string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $EnvironmentName = Get-AdminPowerAppEnvironment -Default | Select -Expand EnvironmentName
            }

            if(-not [string]::IsNullOrWhiteSpace($EnvironmentName))
            {
                $getAllConnectorsRoute = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/apis?showApisWithToS=true&api-version={apiVersion}&`$expand=permissions(`$filter=maxAssignedTo(%27{userId}%27))&`$filter=environment%20eq%20`'{environment}`'" `
                | ReplaceMacro -Macro "{userId}" -Value $Global:currentSession.userId `
                | ReplaceMacro -Macro "{apiVersion}" -Value $ApiVersion `
                | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

                $hbiApis = InvokeApi -Method GET -Route $getAllConnectorsRoute -ApiVersion $ApiVersion | Select -Expand value | `
                %{ New-Object -TypeName PSObject -Property @{ id = $_.id; name = $_.properties.displayName; type = $_.type } }
            }

            $DisplayName = $DisplayName + " (PREVIEW: BlockNonBusinessDataGroup)"
            $lbiDescription = $lbiDescription + " (Blocked)"
            $rules.Add("apiGroupRule", @{
                            actions = @{
                                blockAction = @{
                                    type = "Block"
                                }
                            }
                            parameters = @{
                                apiGroup = "lbi"
                            }
                            type = "ApiGroupRestriction"
                        })
        }

        $CreatedTime = Get-Date -Format "o"

        $newPolicy = @{
            id = ""
            name = ""
            type = $type
            tags = @{}
            properties = @{
                createdTime = $CreatedTime
                displayName = $DisplayName
                definition = @{
                    "`$schema" = $schema
                    defaultApiGroup = "lbi"
                    constraints = $constraints
                    apiGroups = @{
                        hbi = @{
                            apis = $hbiApis
                            description = "Business data only"
                        }
                        lbi = @{
                            apis =  @()
                            description = $lbiDescription
                        }
                    }
                    rules = $rules
                }
            }
        }

        $response = InvokeApi -Method POST -Route $createApiPolicyRoute -Body $newPolicy -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateApiPolicyObject -PolicyObject $response
    }
}

function Remove-AdminDlpPolicy
{
    <#
    .SYNOPSIS
    Deletes the specific Api policy. Delete is successful if it returns a 202 response, 204 means it did not delete.
    .DESCRIPTION
    Remove-AdminDlpPolicy cmdlet deletes a DLP policy.
    Use Get-Help Remove-AdminDlpPolicy -Examples for more detail.
    .PARAMETER PolicyName
    Finds policy matching the specified name.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Remove-AdminDlpPolicy -PolicyName 8c02d657-ad72-4bb9-97c5-afedc4bcf24b
    Deletes policy 8c02d657-ad72-4bb9-97c5-afedc4bcf24b from tenant.
    .EXAMPLE
    Remove-AdminDlpPolicy -EnvironmentName Default-02c201b0-db76-4a6a-b3e1-a69202b479e6 -PolicyName 8c02d657-ad72-4bb9-97c5-afedc4bcf24b
    Deletes policy 8c02d657-ad72-4bb9-97c5-afedc4bcf24b from environment Default-02c201b0-db76-4a6a-b3e1-a69202b479e6.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        if([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies/{policyName}?api-version={apiVersion}"
        }
        else
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/apiPolicies/{policyName}?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
        }

        $route = $route | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Set-AdminDlpPolicy
{
    <#
    .SYNOPSIS
    Updates a policy's environment and default api group settings. Upserts the environment list input (does not append).
    .DESCRIPTION
    Set-AdminDlpPolicy cmdlet updates details on the policy, such as environment filter and default api group.
    Use Get-Help Set-AdminDlpPolicy -Examples for more detail.
    .PARAMETER PolicyName
    Policy name that will be updated.
    .PARAMETER FilterType
    Identifies which filter type the policy will have, none, include or exclude.
    .PARAMETER Environments
    Comma seperated string list used as input environments to either include or exclude, depending on the FilterType.
    .PARAMETER DefaultGroup
    The default group setting, hbi or lbi.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER SetNonBusinessDataGroupState
    Set non business data group(lbi) to Block or Unblock state.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .PARAMETER SchemaVersion
    Specifies the schema version to use, 2016-11-01-preview or 2018-11-01 (HTTP connectors included).
    .EXAMPLE
    Set-AdminDlpPolicy -PolicyName 78d6c98c-aaa0-4b2b-91c3-83d211754d8a -FilterType None
    Clears the environment filter for the policy 78d6c98c-aaa0-4b2b-91c3-83d211754d8a.
    .EXAMPLE
    Set-AdminDlpPolicy -PolicyName 78d6c98c-aaa0-4b2b-91c3-83d211754d8a -FilterType Include -Environments "febb5387-84d7-4717-8345-334a34402f3d,83d98843-bfd7-47ef-bfcd-dc628810ae7b"
    Only applies the policy to the environments febb5387-84d7-4717-8345-334a34402f3d and 83d98843-bfd7-47ef-bfcd-dc628810ae7b.
    .EXAMPLE
    Set-AdminDlpPolicy -PolicyName 78d6c98c-aaa0-4b2b-91c3-83d211754d8a -FilterType Exclude -Environments "febb5387-84d7-4717-8345-334a34402f3d,83d98843-bfd7-47ef-bfcd-dc628810ae7b"
    Applies the policy to all environments except febb5387-84d7-4717-8345-334a34402f3d and 83d98843-bfd7-47ef-bfcd-dc628810ae7b.
    .EXAMPLE
    Set-AdminDlpPolicy -PolicyName 78d6c98c-aaa0-4b2b-91c3-83d211754d8a -DefaultGroup hbi
    Sets the default data group attribute to be hbi (Business data only).
    .EXAMPLE
    Set-AdminDlpPolicy -PolicyName 78d6c98c-aaa0-4b2b-91c3-83d211754d8a -SchemaVersion 2018-11-01
    Sets the DLP Policy to schema version '2018-11-01', allowing for the use of HTTP connectors.
    .EXAMPLE
    Set-AdminDlpPolicy -PolicyName 78d6c98c-aaa0-4b2b-91c3-83d211754d8a -SetNonBusinessDataGroupState "Block"
    Sets non business data(lbi) to block state.
    #>
    [CmdletBinding(DefaultParameterSetName="TenantPolicy")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "TenantPolicy")]
        [Parameter(Mandatory = $true, ParameterSetName = "EnvironmentPolicy")]
        [string]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "TenantPolicy")]
        [ValidateSet("None","Include","Exclude")][string]$FilterType,

        [Parameter(Mandatory = $false, ParameterSetName = "TenantPolicy")]
        [string]$Environments,

        [Parameter(Mandatory = $false, ParameterSetName = "TenantPolicy")]
        [Parameter(Mandatory = $false, ParameterSetName = "EnvironmentPolicy")]
        [ValidateSet("hbi","lbi")][string]$DefaultGroup,

        [Parameter(Mandatory = $true, ParameterSetName = "EnvironmentPolicy")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "TenantPolicy")]
        [Parameter(Mandatory = $false, ParameterSetName = "EnvironmentPolicy")]
        [ValidateSet("Block","Unblock")][string]$SetNonBusinessDataGroupState,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01",

        [Parameter(Mandatory = $false)]
        [ValidateSet("2016-11-01-preview","2018-11-01")][string]$SchemaVersion
    )
    process
    {
        if([string]::IsNullOrWhiteSpace($EnvironmentName))
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies/{policyName}?api-version={apiVersion}"
        }
        else
        {
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/apiPolicies/{policyName}?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName
            $FilterType = "Include"
            $Environments = $EnvironmentName
        }

        $route = $route | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $policy = InvokeApi -Route $route -Method GET -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($FilterType -eq "None")
        {
            $policy.properties.definition.constraints = @{}
        }
        elseif (-not [string]::IsNullOrWhiteSpace($FilterType))
        {
            if ([string]::IsNullOrWhiteSpace($Environments))
            {
                Write-Error "Environments parameter cannot be empty if assigning included or excluded environments to a policy"
                return
            }

            $getEnvironmentUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}?`$expand=permissions&api-version={apiVersion}";

            $environmentInput = (($Environments -replace "` ","") -split ",") | %{ InvokeApi -Method GET -Route ($getEnvironmentUri | ReplaceMacro -Macro "{environmentName}" -Value $_) -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) } `
            | %{ New-Object -TypeName PSObject -Property @{ id = $_.id; name = $_.name; type = $_.type } }

            $constraints = @{
                environmentFilter1 = @{
                    parameters = @{
                        environments = @($environmentInput)
                        filterType = $FilterType.toLower()
                    }
                    type = "environmentFilter"
                }
            }

            $policy.properties.definition.constraints = $constraints
        }

        if (-not [string]::IsNullOrWhiteSpace($DefaultGroup))
        {
            $policy.properties.definition.defaultApiGroup = $DefaultGroup
        }
        if (-not [string]::IsNullOrWhiteSpace($SchemaVersion))
        {
            $schema = "https://schema.management.azure.com/providers/Microsoft.BusinessAppPlatform/schemas/{schemaVersion}/apiPolicyDefinition.json#" `
            | ReplaceMacro -Macro "{schemaVersion}" -Value $SchemaVersion;

            $policy.properties.definition."`$schema" = $schema
        }
        if (-not [string]::IsNullOrWhiteSpace($SetNonBusinessDataGroupState))
        {
            if([bool]($policy.properties.definition.rules.PSobject.Properties.name -match "apiGroupRule") -eq $false)
            {
                $policy.properties.definition.rules | Add-Member -NotePropertyName apiGroupRule -NotePropertyValue $null
            }

            if ($SetNonBusinessDataGroupState -eq "Block")
            {
                $apiGroupRule = @{
                                actions = @{
                                    blockAction = @{
                                        type = "Block"
                                    }
                                }
                                parameters = @{
                                    apiGroup = "lbi"
                                }
                                type = "ApiGroupRestriction"
                            }

                $policy.properties.definition.rules.apiGroupRule = $apiGroupRule

                $policy.properties.definition.apiGroups.lbi.description = "No business data allowed (Blocked)"
            }
            else
            {
                $policy.properties.definition.rules.PSObject.Properties.Remove("apiGroupRule")

                $policy.properties.definition.apiGroups.lbi.description = "No business data allowed"
            }
        }
        $response = InvokeApi -Method PUT -Route $route -Body $policy -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-DlpPolicy
{
    <#
    .SYNOPSIS
    Retrieves a list of DLP policy objects.
    .DESCRIPTION
    Get-DlpPolicy cmdlet gets policy objects for the logged in admin's tenant.
    Use Get-Help Get-DlpPolicy -Examples for more detail.
    .PARAMETER PolicyName
    Get the specific policy by using policy name.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-DlpPolicy
    Retrieves all policies in the tenant.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        if(-not [string]::IsNullOrEmpty($PolicyName))
        {
            # get a policy by policy name
            $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/policies/{policyName}" `
            | ReplaceMacro -Macro "{policyName}" -Value $PolicyName
        }
        else
        {
            # get all policies
            $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/policies?`$top=50"
        }

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-DlpPolicy
{
    <#
    .SYNOPSIS
    Creates a new DLP policy in the tenant by using NewPolicy DLPPolicyDefinition object
    .DESCRIPTION
    New-DlpPolicy cmdlet creates a new DLP policy for the logged in admin's tenant.
    Use Get-Help New-DlpPolicy -Examples for more detail.
    .PARAMETER NewPolicy
    Creates a DLP policy with NewPolicy object.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    New-DlpPolicy -NewPolicy $NewPolicy
    Creates a new policy with $NewPolicy object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Policy")]
        [object]$NewPolicy,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$DisplayName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$EnvironmentType,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [String]$DefaultConnectorClassification = "General",

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [object]$Environments = @(),

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $createApiPolicyRoute = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/policies"

        if ($NewPolicy -eq $null)
        {
            $newPolicy = [pscustomobject]@{
                displayName = $DisplayName
                defaultConnectorsClassification = $DefaultConnectorClassification
                connectorGroups = @()
                environmentType = $EnvironmentType
                environments = $Environments
                etag = $null
            }
        }

        return InvokeApi -Method POST -Route $createApiPolicyRoute -Body $NewPolicy -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-DlpPolicy
{
    <#
    .SYNOPSIS
    Deletes the specific DLP policy by PolicyName.
    .DESCRIPTION
    Remove-DlpPolicy cmdlet deletes a DLP policy.
    Use Get-Help Remove-DlpPolicy -Examples for more detail.
    .PARAMETER PolicyName
    The policy with PolicyName will be deleted.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Remove-DlpPolicy -PolicyName "test policy"
    Deletes policy "test policy" from tenant.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [object]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/policies/{policyName}" `
        | ReplaceMacro -Macro "{policyName}" -Value $PolicyName

        $response = InvokeApi -Method DELETE -Route $route -Body $PolicyToDelete -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Set-DlpPolicy
{
    <#
    .SYNOPSIS
    Updates a policy by using UpdatedPolicy DLPPolicyDefinition object.
    .DESCRIPTION
    Set-DlpPolicy cmdlet updates details on the policy, such as policy display name.
    Use Get-Help Set-DlpPolicy -Examples for more detail.
    .PARAMETER PolicyName
    The policy with PolicyName will be updated.
    .PARAMETER UpdatedPolicy
    Policy that will be updated.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-DlpPolicy -UpdatedPolicy $UpdatedPolicy
    Update the policy to $UpdatedPolicy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$UpdatedPolicy,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/policies/{policyName}" `
        | ReplaceMacro -Macro "{policyName}" -Value $PolicyName

        $response = InvokeApi -Method PATCH -Route $route -Body $UpdatedPolicy -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-PowerAppDlpPolicyConnectorConfigurations
{
    <#
    .SYNOPSIS
    Retrieves connector configurations for a DLP policy.
    .DESCRIPTION
    Get-PowerAppDlpPolicyConnectorConfigurations cmdlet gets connector configurations object for the logged in admin's policy.
    Use Get-Help Get-PowerAppDlpPolicyConnectorConfigurations -Examples for more detail.
    .PARAMETER TenantId
    Get the specific connector configurations by using tenant id.
    .PARAMETER PolicyName
    Get the specific connector configurations by using policy name.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppDlpPolicyConnectorConfigurations -TenantId $TenantId -PolicyName $PolicyName
    Retrieves connector configurations for a DLP policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
         # get connector configurations by tenant id and policy name
            $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/policyconnectorconfigurations" `
            | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
            | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppDlpPolicyConnectorConfigurations
{
    <#
    .SYNOPSIS
    Creates new connector configurations for a DLP policy by using NewDlpPolicyConnectorConfigurations DlpPolicyConnectorConfigurationsDefinition object
    .DESCRIPTION
    New-PowerAppDlpPolicyConnectorConfigurations cmdlet creates new connector configurations for the logged in admin's DLP policy.
    Use Get-Help New-PowerAppDlpPolicyConnectorConfigurations -Examples for more detail.
    .PARAMETER TenantId
    Create new connector configurations by using tenant id.
    .PARAMETER PolicyName
    Create new connector configurations by using policy name.
    .PARAMETER NewDlpPolicyConnectorConfigurations
    Creates a DLP policy with NewDlpPolicyConnectorConfigurations object.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    New-PowerAppDlpPolicyConnectorConfigurations -TenantId $TenantId -PolicyName $PolicyName -NewDlpPolicyConnectorConfigurations $NewDlpPolicyConnectorConfigurations
    Creates a new policy with $TenantId, $PolicyName and $NewDlpPolicyConnectorConfigurations object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$NewDlpPolicyConnectorConfigurations,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/policyconnectorconfigurations" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        return InvokeApi -Method POST -Route $route -Body $NewDlpPolicyConnectorConfigurations -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-PowerAppDlpPolicyConnectorConfigurations
{
    <#
    .SYNOPSIS
    Deletes the specific DLP policy connector configurations by TenantId and PolicyName.
    .DESCRIPTION
    Remove-PowerAppDlpPolicyConnectorConfigurations cmdlet deletes connector configurations for a DLP policy.
    Use Get-Help Remove-PowerAppDlpPolicyConnectorConfigurations -Examples for more detail.
    .PARAMETER TenantId
    The policy with TenantId will be deleted.
    .PARAMETER PolicyName
    The policy with PolicyName will be deleted.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Remove-PowerAppDlpPolicyConnectorConfigurations -TenantId $TenantId -PolicyName $PolicyName
    Deletes connector configurations for a specific DLP policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [String]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/policyconnectorconfigurations" `
            | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
            | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Set-PowerAppDlpPolicyConnectorConfigurations
{
    <#
    .SYNOPSIS
    Updates connector configurations for a DLP policy by using UpdatedConnectorConfigurations DlpPolicyConnectorConfigurationsDefinition object.
    .DESCRIPTION
    Set-PowerAppDlpPolicyConnectorConfigurations cmdlet updates details on the policy connector configurations, such as connector endpoint.
    Use Get-Help Set-PowerAppDlpPolicyConnectorConfigurations -Examples for more detail.
    .PARAMETER TenantId
    The connector configurations with TenantId will be updated.
    .PARAMETER PolicyName
    The connector configurations with PolicyName will be updated.
    .PARAMETER UpdatedConnectorConfigurations
    Policy connector configurations that will be updated.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppDlpPolicyConnectorConfigurations -UpdatedConnectorConfigurations $UpdatedConnectorConfigurations
    Update the DLP policy connector configurations to $UpdatedConnectorConfigurations.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$UpdatedConnectorConfigurations,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/policyconnectorconfigurations" `
            | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
            | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $response = InvokeApi -Method PATCH -Route $route -Body $UpdatedConnectorConfigurations -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Add-AdminPowerAppsSyncUser
{
    <#
    .SYNOPSIS
    Adds a user to the CRM database.
    .DESCRIPTION
    Add-AdminPowerAppsSyncUser cmdlet adds a user to the CRM database.
    Use Get-Help Add-AdminPowerAppsSyncUser -Examples for more detail.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER PrincipalObjectId
    The objectId of a user to be added.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Add-AdminPowerAppsSyncUser -EnvironmentName 8c02d657-ad72-4bb9-97c5-afedc4bcf24b -PrincipalObjectId 24a5b286-2abe-46b4-9ba5-9de73169ab9c
    Add the user (24a5b286-2abe-46b4-9ba5-9de73169ab9c) to the CRM database which is linked in environment 8c02d657-ad72-4bb9-97c5-afedc4bcf24b.
    #>
    [CmdletBinding(DefaultParameterSetName="User")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "User", ValueFromPipelineByPropertyName = $true)]
        [string]$PrincipalObjectId,

        [Parameter(Mandatory = $false, ParameterSetName = "User")]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentName}/addUser?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName

        $userToadd = @{
            ObjectId = $PrincipalObjectId
        }

        $response = InvokeApi -Method POST -Route $route -Body $userToadd -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-AdminPowerAppCdsAdditionalNotificationEmails
{
    <#
    .SYNOPSIS
    Returns email addresses of users other than default admins of CDS that receive notifications.
    .DESCRIPTION
    The Get-AdminPowerAppCdsAdditionalNotificationEmails cmdlet returns email addresses of users other than default admins of CDS linked to an environment that receive notifications.
    Use Get-Help Get-AdminPowerAppCdsAdditionalNotificationEmails -Examples for more details.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-AdminPowerAppCdsAdditionalNotificationEmails -EnvironmentName 02c201b0-db76-4a6a-b3e1-a69202b479e6
    Returns email addresses of users other than default admins of CDS linked to the environment 02c201b0-db76-4a6a-b3e1-a69202b479e6.
    #>
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-10-01"
    )

    $getAdditionalNotificationEmailsUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environmentName}?`$expand=properties.linkedEnvironmentMetadata.additionalNotificationEmails&api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

    $additionalNotificationEmailsResponse = InvokeApiNoParseContent -Method GET -Route $getAdditionalNotificationEmailsUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    if($additionalNotificationEmailsResponse.StatusCode -ne 200)
    {
        CreateHttpResponse($additionalNotificationEmailsResponse)
    }
    else
    {
        $responseJson = ConvertFrom-JsonWithErrorHandling($additionalNotificationEmailsResponse.Content)
        return $responseJson.properties.linkedEnvironmentMetadata.additionalNotificationEmails
    }
}

function Set-AdminPowerAppCdsAdditionalNotificationEmails
{
    <#
    .SYNOPSIS
    Sets email addresses of users other than default admins of CDS that should receive notifications.
    .DESCRIPTION
    The Set-AdminPowerAppCdsAdditionalNotificationEmails cmdlet sets email addresses of users other than default admins of CDS linked to an environment that should receive notifications.
    Use Get-Help Set-AdminPowerAppCdsAdditionalNotificationEmails -Examples for more details.
    .PARAMETER EnvironmentName
    The Environment's identifier.
    .PARAMETER AdditionalNotificationEmails
    The email addresses
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-AdminPowerAppCdsAdditionalNotificationEmails -EnvironmentName 02c201b0-db76-4a6a-b3e1-a69202b479e6 -AdditionalNotificationEmails abc@test.com
    Returns email addresses of users other than default admins of CDS linked to the environment 02c201b0-db76-4a6a-b3e1-a69202b479e6.
    #>
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$AdditionalNotificationEmails,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-10-01"
    )

    $setAdditionalNotificationEmailsUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environmentName}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

    $requestBody = @{
            properties = @{
                linkedEnvironmentMetadata = @{
                    additionalNotificationEmails = $AdditionalNotificationEmails
                }
            }
        }

    $setAdditionalNotificationEmailsResponse = InvokeApi -Method PATCH -Route $setAdditionalNotificationEmailsUri -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    if($setAdditionalNotificationEmailsResponse.StatusCode -ne 202)
    {
        Write-Error "Operation to set email addresses failed!"
    }

    CreateHttpResponse($setAdditionalNotificationEmailsResponse)
}

function Get-AdminPowerAppLicenses
{
    <#
    .SYNOPSIS
    Downloads the user licenses into a specified file.
    .DESCRIPTION
    The Get-AdminPowerAppLicenses cmdlet downloads the user licenses into a specified file.
    Use Get-Help Get-AdminPowerAppLicenses -Examples for more details.
    .PARAMETER OutputFilePath
    The output file path
    .PARAMETER TimeoutInMinutes
    The timeout setting in minutes.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-AdminPowerAppLicenses -OutputFilePath C:\Users\testuser\licenses.json
    Donloads the licenses of calling user into specified path file C:\Users\testuser\licenses.json
    #>
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath = $null,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )

    $getLicensesUri = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/exportServicePlans?api-version={apiVersion}"

    $getLicensesResponse = InvokeApi -Method POST -Route $getLicensesUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    if ($getLicensesResponse.StatusCode -eq 202)
    {
        $getLicensesUri = $getLicensesResponse.Headers['Location']
        $currentTime = Get-Date -format HH:mm:ss
        $nextTime = Get-Date -format HH:mm:ss
        $TimeDiff = New-TimeSpan $currentTime $nextTime

        #Wait until the operation complete, there is an error, or we hit a timeout
        while((-not [string]::IsNullOrEmpty($getLicensesUri)) -and ($getLicensesResponse.StatusCode -eq 202) -and ($TimeDiff.TotalMinutes -lt $TimeoutInMinutes))
        {
            Start-Sleep -s 5
            $getLicensesResponse = InvokeApiNoParseContent -Route $getLicensesUri -Method GET -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            $nextTime = Get-Date -format HH:mm:ss
            $TimeDiff = New-TimeSpan $currentTime $nextTime
        }

        if ($TimeDiff.TotalMinutes -ge $TimeoutInMinutes)
        {
            Write-Error "Get-AdminPowerAppLicenses timeout."
            throw
        }
    }

    if ($getLicensesResponse.Content -ne $null)
    {
        $jobject = ConvertFrom-Json($getLicensesResponse.Content)
        $csvFileUri = $jobject.sharedAccessSignature
        Invoke-WebRequest $csvFileUri -OutFile $OutputFilePath
    }

    CreateHttpResponse($getLicensesResponse)
}

function Set-AdminPowerAppDesiredLogicalName
{
    <#
     .SYNOPSIS
     Sets the desired logical name of the power app.
     .DESCRIPTION
     The Set-AdminPowerAppDesiredLogicalName sets the desired logical name for the specific power app.
     .PARAMETER EnvironmentName
     The environment to which the app belongs.
     .PARAMETER AppName
     App name for the one which you want to set desired logical name.
     .PARAMETER DesiredLogicalName
     The desired logical name to set to the non-solution aware app.
     .EXAMPLE
     Set-AdminPowerAppDesiredLogicalName -EnvironmentName 08b4e32a-4e0d-4a69-97da-e1640f0eb7b9 -AppName 3c2f7648-ad60-4871-91cb-b77d7ef3c239 -DesiredLogicalName pa232_myapp_234232
     Sets the desired logical name with 'pa232_myapp_234232' of app '3c2f7648-ad60-4871-91cb-b77d7ef3c239' in environment '08b4e32a-4e0d-4a69-97da-e1640f0eb7b9'.
     #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$DesiredLogicalName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2020-06-01"
    )

    process
    {
        $setDesiredLogicalNameUri = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environmentname}/apps/{application}/setDesiredLogicalName?api-version={apiVersion}"`
        | ReplaceMacro -Macro "{environmentname}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{application}" -Value $AppName;

        #Construct the request body
        $setDesiredLogicalNameRequestBody = @{
            desiredLogicalName = $DesiredLogicalName
        }

        $response = InvokeApi -Method POST -Route $setDesiredLogicalNameUri -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true) -Body $setDesiredLogicalNameRequestBody

        CreateHttpResponse($response)
    }
}

function Get-AdminDeletedPowerAppsList
{
    <#
    .SYNOPSIS
    Returns the list of deleted power apps in the admin's specified environment.
    .DESCRIPTION
    The Get-AdminDeletedPowerAppsList function returns all deleted power apps in the given environment.
    Use Get-Help Get-AdminDeletedPowerAppsList -Examples for more detail.
    .PARAMETER EnvironmentName
    The environment for the deleted power apps.
    .EXAMPLE
    Get-AdminDeletedPowerAppsList -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f
    Returns all deleted power apps in the admin's specified environment with name 0fc02431-15fb-4563-a5ab-8211beb2a86f.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2017-06-01"
    )
    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/deletedApps?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $deletedPowerAppsResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion

        if ($deletedPowerAppsResult.StatusCode -eq "NotFound")
        {
            Write-Error "The specified power app environment was not found."
            return
        }
        else
        {
            foreach ($deletedApp in $deletedPowerAppsResult.Value)
            {
                CreateAppObject -AppObj $deletedApp;
            }
        }
    }
}

function Get-AdminRecoverDeletedPowerApp
{
    <#
    .SYNOPSIS
    Recovers the deleted power app with the specified app ID in the specified environment.
    Please note that after the power app is recovered, the administrator needs to manually recreate the permissions for this app. 
    .DESCRIPTION
    The Get-AdminRecoverDeletedPowerApp function recovers the deleted power app with the specified app ID in the specified environment.
    Use Get-Help Get-AdminRecoverDeletedPowerApp -Examples for more detail.
    .PARAMETER AppName
    The app id/name of the deleted power app.
    .PARAMETER EnvironmentName
    The environment for the deleted power app.
    .EXAMPLE
    Get-AdminRecoverDeletedPowerApp -AppName 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488 -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f
    Recovers the deleted app with name 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488 in the admin's specified environment with name 0fc02431-15fb-4563-a5ab-8211beb2a86f.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/deletedApps/{appName}/restore?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{appName}" -Value $AppName `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $recoveredPowerAppsResult = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion

        if ($recoveredPowerAppsResult.StatusCode -eq "NotFound")
        {
            Write-Error "The specified power app was not found."
            return
        }
        else
        {
            foreach ($recoveredPowerApp in $recoveredPowerAppsResult.Value)
            {
                CreateAppObject -AppObj $recoveredPowerApp;
            }
            CreateHttpResponse($recoveredPowerAppsResult);
        }
    }
}

function Add-AdminAllowedThirdPartyApps
{
    <#
    .SYNOPSIS
    Adds the application to the list of allowed third party applications for the tenant.
    .DESCRIPTION
    The Add-AdminAllowedThirdPartyApps function adds the allowed third party app ID.
    Use Get-Help Add-AdminAllowedThirdPartyApps -Examples for more detail.
    .PARAMETER ApplicationId
    The app id/name of the allowed third party app.
    .EXAMPLE
    Add-AdminAllowedThirdPartyApps -ApplicationId 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488
    Adds the application Id to the list of allowed third party applications for the tenant.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ApplicationId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/allowedThirdPartyApps/{applicationId}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{applicationId}" -Value $ApplicationId `

        return InvokeApi -Method PUT -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-AdminAllowedThirdPartyApps
{
    <#
    .SYNOPSIS
    Adds the application to the list of allowed third party applications for the tenant.
    .DESCRIPTION
    The Get-AdminAllowedThirdPartyApps function retrieves the list of allowed third party app ids.
    Use Get-Help Get-AdminAllowedThirdPartyApps -Examples for more detail.
    .EXAMPLE
    Get-AdminAllowedThirdPartyApps
    Adds the application Id to the list of allowed third party applications for the tenant.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/allowedThirdPartyApps?api-version={apiVersion}"

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion
    }
}

function Remove-AdminAllowedThirdPartyApps
{
    <#
    .SYNOPSIS
    Removes the application from the list of allowed third party applications for the tenant.
    .DESCRIPTION
    The Remove-AdminAllowedThirdPartyApps function removes the allowed third party app ids.
    Use Get-Help Remove-AdminAllowedThirdPartyApps -Examples for more detail.
    .PARAMETER ApplicationId
    The app id/name of the allowed third party app that will be removed.
    .EXAMPLE
    Remove-AdminAllowedThirdPartyApps -ApplicationId 1ec3c80c-c2c0-4ea6-97a8-31d8c8c3d488
    Adds the application Id to the list of allowed third party applications for the tenant.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ApplicationId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/allowedThirdPartyApps/{applicationId}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{applicationId}" -Value $ApplicationId `

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}


function Start-DLPEnforcementOnConnectionsInTenant
{
    <#
    .SYNOPSIS
    Enforce DLP policies on connections in the tenant. This will disable connections that are violating DLP policies.
    .DESCRIPTION
    The Start-DLPEnforcementOnConnectionsInTenant function will disable connections that are violating DLP policies in the tenant.
    Use Get-Help Start-DLPEnforcementOnConnectionsInTenant -Examples for more detail.
    .EXAMPLE
    Start-DLPEnforcementOnConnectionsInTenant 
    Enforce DLP policies on connections in the tenant. This will disable connections that are violating DLP policies.
    #>
    [CmdletBinding()]
    param
    (
        
        [Parameter(Mandatory = $false)]
        [string]$JobMode = "EvaluateAndUpdateConnections",

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
			jobMode = $JobMode
		};

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/reevaluateconnectionsdlp?api-version={apiVersion}"
        $response = InvokeApiNoParseContent -Method POST -Body $requestBody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($response.StatusCode -eq "202")
        {
            $successMessage = 'A background job to enforce Data Loss Prevention (DLP) policies in connections within this tenant has been created and is now running. Please note that this process may take a few hours to complete.'
            CreateHttpResponse($successMessage)
        }
        else
        {
            CreateHttpResponse($response)
        }
    }
}

function Start-DLPEnforcementOnConnectionsInEnvironment
{
    <#
    .SYNOPSIS
    Enforce DLP policies on connections in an environment. This will disable connections that are violating DLP policies.
    .DESCRIPTION
    The Start-DLPEnforcementOnConnectionsInEnvironment function will disable connections that are violating DLP policies in an environment.
    Use Get-Help Start-DLPEnforcementOnConnectionsInEnvironment -Examples for more detail.
    .EXAMPLE
    Start-DLPEnforcementOnConnectionsInEnvironment
    Start-DLPEnforcementOnConnectionsInEnvironment -EnvironmentName c4a07cd6-cb14-e987-b5a2-a1dd61346963 
    Enforce DLP policies on connections in an environment. This will disable connections that are violating DLP policies.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$JobMode = "EvaluateAndUpdateConnections",

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-02-01"
    )
    process
    {
        $requestBody = @{
			jobMode = $JobMode
		};

        $route = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/scopes/admin/environments/{environment}/reevaluateconnectionsdlp?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $response = InvokeApiNoParseContent -Method POST -Body $RequestBody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($response.StatusCode -eq "202")
        {
            $successMessage = 'A background job to enforce Data Loss Prevention (DLP) policies in connections within this environment has been created and is now running. Please note that this process may take a half an hour to complete.'
            CreateHttpResponse($successMessage)
        }
        else
        {
            CreateHttpResponse($response)
        }
    }
}

function Copy-PowerAppEnvironment
{
    <#
    .SYNOPSIS
    Copy an environment from source to target.
    .DESCRIPTION
    The Copy-PowerAppEnvironment function copies an environment from source to target.
    Use Get-Help Copy-PowerAppEnvironment -Examples for more detail.
    .PARAMETER EnvironmentName
    The target environment name.
    .PARAMETER CopyToRequestDefinition
    The copy request definition object.
    .PARAMETER WaitUntilFinished
    If set to true, the function will not return until complete.
    .PARAMETER TimeoutInMinutes
    The timeout setting in minutes.
    .EXAMPLE
    Copy-PowerAppEnvironment -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -CopyToRequestDefinition $copyToRequest
    Copy a source environment to 0fc02431-15fb-4563-a5ab-8211beb2a86f.
        $copyToRequest = [pscustomobject]@{
            SourceEnvironmentId = $sourceEnvironment.EnvironmentName
            TargetEnvironmentName = "Copied from source"
            TargetSecurityGroupId = "204162d5-59db-40c2-9788-2cda6b063f2b"
            CopyType = "MinimalCopy"
            SkipAuditData: true
        }
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [object]$CopyToRequestDefinition,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $false,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/copyTo?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $response = InvokeApi -Method POST -Body $CopyToRequestDefinition -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        # Poll until the copy is completed
        If($WaitUntilFinished)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function TenantToTenant-PrepareMigration
{
    <#
    .SYNOPSIS
    Validate a Tenant To Tenant Migration of an environment from source tenant to target tenant.
    .DESCRIPTION
    The TenantToTenant-PrepareMigration function validates the migration of an environment from source tenant to target tenant.
    Use Get-Help TenantToTenant-PrepareMigration -Examples for more detail.
    .PARAMETER MigrationId
    The Migration ID provided in the TenantToTenant-ViewMigrationRequest.
    .PARAMETER TargetTenantId
    The TargetTenantId for the Migration.
    .PARAMETER ReadOnlyUserMappingFileContainerUri
    The ReadOnlyUserMappingFileContainerUri for the Migration. Perform GenerateResourceStorage-PowerAppEnvironment to get the container URL for the environment and upload the usermapping file to the path. This Uri is the container Uri where you have uploaded the usermapping file.
    .PARAMETER SecurityGroupId
    The Azure Active Directory security group object identifier to restrict database membership.
    .PARAMETER WaitUntilFinished
    If set to true, the function will not return until complete validation of Tenant To Tenant Migration is complete.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    
    .EXAMPLE
    TenantToTenant-PrepareMigration -MigrationId 0fc02431-15fb-4563-a5ab-8211beb2a86f -TargetTenantId 26fc7803-4578-4760-as23-f0afb76b84cc -ReadOnlyUserMappingFileContainerUri https://contoso.blob.core.windows.net/readonlyusermappingfilecontainer
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$MigrationId,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$TargetTenantId,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ReadOnlyUserMappingFileContainerUri,

        [Parameter(Mandatory = $false)]
        [string]$SecurityGroupId,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $false,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )
    DynamicParam 
    {
            # Display warning message
            Write-Warning "Kindly consult the provided document for details on the components. Some components may not be fully supported and could necessitate additional actions on your part or require assistance from our support team."
            Write-Warning "https://learn.microsoft.com/en-us/power-platform/admin/move-environment-tenant?tabs=image#confirm-if-any-of-the-solutions-below-are-installed-in-the-environments-to-be-migrated-as-these-may-require-more-steps-either-from-you-or-support"

            # Return an empty dictionary as no dynamic parameters are needed
            $null = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
    }
    process
    {

        # Remove SAS token if present
        $uri = New-Object System.Uri($ReadOnlyUserMappingFileContainerUri)
        $containerUri = $uri.GetLeftPart([System.UriPartial]::Path)

        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrations/{t2tmigrationId}/prepare?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{t2tmigrationId}" -Value $MigrationId;

        $requestBody = @{
                TargetTenantId = $TargetTenantId
                ReadOnlyUserMappingFileContainerUri = $containerUri
                NewSecurityGroupId = $SecurityGroupId
            }

        $response = InvokeApi -Method POST -Body $requestBody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        
        # Poll until the copy is completed
        If($WaitUntilFinished)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function TenantToTenant-GetMigrationStatus
{
    <#
    .SYNOPSIS
    Get Status of a Tenant To Tenant Migration of an Environment.
    .DESCRIPTION
    The TenantToTenant-GetMigrationStatus function validates the status of Tenant To Tenant migration of an environment from source to target.
    Use Get-Help TenantToTenant-GetMigrationStatus -Examples for more detail.
    .PARAMETER MigrationId
    The Migration Id provided in the TenantToTenant-ViewMigrationRequest command.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    
    .EXAMPLE
    TenantToTenant-GetMigrationStatus -MigrationId 0fc02431-15fb-4563-a5ab-8211beb2a86f
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$MigrationId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrations/{migrationId}/get?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{migrationId}" -Value $MigrationId;

        $response = InvokeApi -Method Get -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        
        CreateHttpResponse($response)
    }
}

function TenantToTenant-MigratePowerAppEnvironment
{
    <#
    .SYNOPSIS
    Tenant To Tenant Migration of an environment from source to target. Requires running the Tenant To Tenant Prepare Migration function first to ensure a smooth transition.
    .DESCRIPTION
    The TenantToTenant-MigratePowerAppEnvironment function migrates an environment from source to target tenant.
    Use Get-Help TenantToTenant-MigratePowerAppEnvironment -Examples for more detail.
    .PARAMETER MigrationId
    The Migration ID Provided in TenantToTenant-ViewMigrationRequest.
    .PARAMETER TargetTenantId
    The TargetTenantId for the Migration.
    .PARAMETER SecurityGroupId 
    The Security Group ID for the newly Migrated Environment.
    .PARAMETER WaitUntilFinished
    If set to true, the function will not return until complete Tenant To Tenant Migration is completed.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    
    .EXAMPLE
    TenantToTenant-MigratePowerAppEnvironment -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -TargetTenantId 26fc7803-4578-4760-as23-f0afb76b84cc
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$MigrationId,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$TargetTenantId,
        
        [Parameter(Mandatory = $false)]
        [string]$SecurityGroupId,

        [Parameter(Mandatory=$false)]
        [bool]$WaitUntilFinished = $false,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )
    DynamicParam 
    {
            # Display warning message
            Write-Warning "Kindly consult the provided document for details on the components. Some components may not be fully supported and could necessitate additional actions on your part or require assistance from our support team."
            Write-Warning "https://learn.microsoft.com/en-us/power-platform/admin/move-environment-tenant?tabs=image#confirm-if-any-of-the-solutions-below-are-installed-in-the-environments-to-be-migrated-as-these-may-require-more-steps-either-from-you-or-support"

            # Return an empty dictionary as no dynamic parameters are needed
            $null = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
    }
    process
    {
       $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrations/{t2tmigrationId}/migrate?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{t2tmigrationId}" -Value $MigrationId;

        $requestBody = @{
                TargetTenantId = $TargetTenantId
            }

        if (-not [string]::IsNullOrEmpty($SecurityGroupId))
        {
            $requestBody += @{
                NewSecurityGroupId = $SecurityGroupId
            }
        }

        $response = InvokeApi -Method POST -Body $requestBody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
       
        # Poll until the copy is completed
        If($WaitUntilFinished)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function GenerateResourceStorage-PowerAppEnvironment
{
    <#
    .SYNOPSIS
    Generate Resource Storage for an environment. Can be used to upload Usermapping files for Tenant To Tenant Migration.
    .DESCRIPTION
    The GenerateResourceStorage-PowerAppEnvironment function generates a storage specific to the environment.
    Use Get-Help GenerateResourceStorage-PowerAppEnvironment -Examples for more detail.
    .PARAMETER EnvironmentName
    The Environment name.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    
    .EXAMPLE
    GenerateResourceStorage-PowerAppEnvironment -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory=$false)]
        [bool]$WaitUntilFinished = $false,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/generateTenantToTenantResourceStorage?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        Write-Host "API Route: $route"

        $response = InvokeApi -Method POST -Body $requestBody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
       
        CreateHttpResponse($response)
    }
}

function TenantToTenant-UploadUserMappingFile
{
    <#
    .SYNOPSIS
    Generates resource storage and uploads a user mapping file for a specified environment.
    .DESCRIPTION
    This function combines the generation of resource storage and the upload of a user mapping file into a single step.
    .PARAMETER EnvironmentName
    The ID of the PowerApps environment.
    .PARAMETER UserMappingFilePath
    The full path to the user mapping file (e.g., C:\path\to\usermapping.csv).
    .PARAMETER ApiVersion
    The API version to call with. Default is 2021-04-01.
    .EXAMPLE
    TenantToTenant-UploadUserMappingFile -EnvironmentName "0fc02431-15fb-4563-a5ab-8211beb2a86f" -UserMappingFilePath "C:\path\to\usermapping.csv"
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true)]
        [string]$UserMappingFilePath,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        try {
            Write-Host "Checking Az.Storage module installation"
            # Check if Az module is installed
            if (-not (Get-Module -ListAvailable -Name Az.Storage)) {
                Write-Host "Az module not found. Installing Az module..."
                Install-Module -Name  Az.Storage -Repository PSGallery -Force -AllowClobber -Scope CurrentUser -ErrorAction Continue
            }

            Write-Host "Importing Az.Storage Module"

            # Import Az module
            Import-Module Az.Storage -ErrorAction stop

            Write-Host "Generating resource storage for environment: $EnvironmentName"

            # Generate Resource Storage
            $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/generateTenantToTenantResourceStorage?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

            $response = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            $sasUri = $response.sharedAccessSignature;

            # Parse SAS URI
            $uri = [System.Uri]$sasUri
            $storageAccountName = $uri.DnsSafeHost.Split(".")[0]
            $container = $uri.LocalPath.TrimStart("/")
            $sasToken = $uri.Query

            # Create a storage context
            $storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -SasToken $sasToken

            # Upload the file to Azure Blob Storage
            Set-AzStorageBlobContent -File $UserMappingFilePath -Container $container -Context $storageContext -Force

            Write-Host "File uploaded successfully to container $container in storage account $storageAccountName`n"
            $baseUri = $uri.GetLeftPart([System.UriPartial]::Path)
            Write-Host "Read Only UserMapping File ContainerUri: $baseUri"
        }
        catch {
                if($response.Error -ne $null)
                {
                    CreateHttpResponse($response.Error)
                }
                else
                {
                    Write-Error "An error occurred: $_"
                }
        }
    }
}


function TenantToTenant-SubmitMigrationRequest
{
    <#
    .SYNOPSIS
    Submit a Tenant To Tenant migration request from Source Tenant to a Destination Tenant.
    .DESCRIPTION
    The TenantToTenant-SubmitMigrationRequest function submits a migration request from the source tenant to the destination tenant.
    Use Get-Help TenantToTenant-SubmitMigrationRequest -Examples for more detail.
    .PARAMETER EnvironmentName
    The Environment name.
    .PARAMETER TargetTenantID
    The target tenant ID.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    .EXAMPLE
    TenantToTenant-SubmitMigrationRequest -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -TargetTenantID 12345678-90ab-cdef-1234-567890abcdef
    Submits a migration request from the source tenant to the destination tenant.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$TargetTenantID,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/createt2t?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $requestBody = @{
                TargetTenantId = $TargetTenantID
            }

        $response = InvokeApi -Method POST -Body $requestBody -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($response -ne $null -and $response.StatusCode -eq 200)
        {
            Write-Output "Approval request sent to destination tenant $TargetTenantID. Please reach out to Target tenant ID admin to approve the request."
        }
        else
        {
            Write-Error "Failed to send approval request."
        }

        CreateHttpResponse($response)
    }
}

function TenantToTenant-DeleteMigrationRequest
{
    <#
    .SYNOPSIS
    Delete a Migration Request using RequestId.
    .DESCRIPTION
    The TenantToTenant-DeleteMigrationRequest function deletes a migration request from the tenant.
    Use Get-Help TenantToTenant-DeleteMigrationRequest -Examples for more detail.
    .PARAMETER MigrationId
    The Migration ID Identifier to delete the request.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    .EXAMPLE
    TenantToTenant-DeleteMigrationRequest -MigrationId 0fc02431-15fb-4563-a5ab-8211beb2a86f
    Deletes a migration request from the tenant.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$MigrationId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrations/{migrationId}/delete?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{migrationId}" -Value $MigrationId;

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function TenantToTenant-ManageMigrationRequest {
    <#
    .SYNOPSIS
    Approving/Rejecting Tenant To Tenant Migration using the Approval Request Sent by the source Tenant Admin.
    .DESCRIPTION
    The TenantToTenant-ManageMigrationRequest function approves/rejects a request from a source tenant for migration of an environment into the current tenant.
    Use Get-Help TenantToTenant-ManageMigrationRequest -Examples for more detail.
    .PARAMETER MigrationId
    The Migration Id.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    .EXAMPLE
    TenantToTenant-ManageMigrationRequest -MigrationId 0fc02431-15fb-4563-a5ab-8211beb2a86f
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$MigrationId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )

    process {

            $requestBody = @{
                migrationId = $MigrationId
            }

            $approvalStatus = Read-Host "Enter approval status for MigrationId $migrationId (0 for Rejected, 1 for Approved):"

            if ($approvalStatus -eq "1") 
                { 
                    $approvalRoute =   "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrationapprovals/{approvalRequestId}/approve?api-version=$ApiVersion" `
                    | ReplaceMacro -Macro "{approvalRequestId}" -Value $migrationId;
                }
            else 
				{
					$approvalRoute = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrationapprovals/{approvalRequestId}/reject?api-version=$ApiVersion" `
                    | ReplaceMacro -Macro "{approvalRequestId}" -Value $requestId;
				}

            $approvalResponse = InvokeApi -Method POST -Body $approvalBody -Route $approvalRoute -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

            CreateHttpResponse($approvalResponse)
        } 
}

function TenantToTenant-ViewMigrationRequest {
    <#
    .SYNOPSIS
    View Tenant To Tenant Migration Requests for the Tenant.
    .DESCRIPTION
    The TenantToTenant-ViewMigrationRequest function lists all the requests in the tenant.
    Use Get-Help TenantToTenant-ViewMigrationRequest -Examples for more detail.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    .EXAMPLE
    TenantToTenant-ViewMigrationRequest
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )

    process {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrations?api-version=$ApiVersion"
        $response = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($response -and $response.value) {
            $response.value | ForEach-Object {
                [PSCustomObject]@{
                    MigrationId       = $_.id
                    EnvironmentId     = $_.environmentId
                    TargetTenantId    = $_.targetTenantId
                    CreatedOnTime     = $_.createdOnTime
                    CreatedById       = $_.createdBy.id
                    CreatedByType     = $_.createdBy.type
                    ApprovalActionedById = $_.approvalActionedBy.id
                    ApprovalActionedByType = $_.approvalActionedBy.type
                    ApprovalStatus    = $_.approvalStatus
                }
            }
        } else {
            Write-Output "No migration requests found."
        }
    }
}

function TenantToTenant-ViewApprovalRequest {
    <#
    .SYNOPSIS
    View Tenant To Tenant Migration Requests for the Tenant.
    .DESCRIPTION
    The TenantToTenant-ViewApprovalRequest function list all the request in the tenant.
    Use Get-Help TenantToTenant-ViewApprovalRequest -Examples for more detail.
    .PARAMETER ApiVersion
    The api version to call with. Default 2021-04-01.
    .EXAMPLE
    TenantToTenant-ViewApprovalRequest
    #>
    [CmdletBinding()]
    param (

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2021-04-01"
    )

    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/t2tmigrationapprovals?api-version=$ApiVersion"

        $response = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        if ($response -and $response.value) {
            $sevenDaysAgo = (Get-Date).AddDays(-7)
            $approvals = $response.value | ForEach-Object {
                if ([DateTime]$_.resourceCreatedTime -le $sevenDaysAgo) {
                    $_.approvalState = "Expired"
                }
                [PSCustomObject]@{
                    TenantId          = $_.tenantId
                    MigrationId       = $_.id
                    CreatedBy         = $_.createdBy
                    CreatedTime       = $_.resourceCreatedTime
                    ApprovalState     = $_.approvalState
                    SourceMigrationId = $_.sourceMigrationId
                    EnvironmentName   = $_.environmentName
                    OrganizationId    = $_.organizationId
                }
            }

            $sortedApprovals = $approvals | Sort-Object -Property CreatedTime -Descending

            $sortedApprovals | ForEach-Object {
                Write-Output $_
            }
        } else {
            Write-Output "No Approval requests found."
        }
    }
}

function Backup-PowerAppEnvironment
{
    <#
    .SYNOPSIS
    Backup an environment.
    .DESCRIPTION
    The Backup-PowerAppEnvironment function backups an environment.
    Use Get-Help Backup-PowerAppEnvironment -Examples for more detail.
    .PARAMETER EnvironmentName
    The target environment name.
    .PARAMETER BackupRequestDefinition
    The backup request definition object.
    .EXAMPLE
    Backup-PowerAppEnvironment -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -BackupRequestDefinition $backupRequest
    Backup environment 0fc02431-15fb-4563-a5ab-8211beb2a86f.
        $backupRequest = [pscustomobject]@{
            Label = "this is a label"
            Notes = "this is a note"
        }
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [object]$BackupRequestDefinition,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/backups?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        InvokeApi -Method POST -Body $BackupRequestDefinition -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppEnvironmentBackups
{
    <#
    .SYNOPSIS
    Get backup environments.
    .DESCRIPTION
    The Get-PowerAppEnvironmentBackups function gets environment backup list.
    Use Get-Help Get-PowerAppEnvironmentBackups -Examples for more detail.
    .PARAMETER EnvironmentName
    The target environment name.
    .EXAMPLE
    Get-PowerAppEnvironmentBackups -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f
    Get backup list for environment 0fc02431-15fb-4563-a5ab-8211beb2a86f.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/backups?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Restore-PowerAppEnvironment
{
    <#
    .SYNOPSIS
    Restores an environment.
    .DESCRIPTION
    The Restore-PowerAppEnvironment function restores an environment.
    Use Get-Help Restore-PowerAppEnvironment -Examples for more detail.
    .PARAMETER EnvironmentName
    The target environment name.
    .PARAMETER RestoreToRequestDefinition
    The restore request definition object.
    .PARAMETER WaitUntilFinished
    If set to true, the function will not return until complete.
    .PARAMETER TimeoutInMinutes
    The timeout setting in minutes.
    .EXAMPLE
    Restore-PowerAppEnvironment -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -RestoreToRequestDefinition $restoreRequest
    Restore environment 0fc02431-15fb-4563-a5ab-8211beb2a86f.
        $restoreRequest = [pscustomobject]@{
            SourceEnvironmentId = 0fc02431-15fb-4563-a5ab-8211beb2a86f
            TargetEnvironmentName = "Restored Environment"
            TargetSecurityGroupId = "204162d5-59db-40c2-9788-2cda6b063f2b"
            RestorePointDateTime = $datetime.ToString("yyyy-MM-dd HH:mm:ss")
            SkipAuditData: true
        }
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [object]$RestoreToRequestDefinition,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/restoreTo?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $response = InvokeApi -Method POST -Body $RestoreToRequestDefinition -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        # Poll until the retore is completed
        If($WaitUntilFinished)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function Reset-PowerAppEnvironment
{
    <#
    .SYNOPSIS
    Reset environment.
    .DESCRIPTION
    The Reset-PowerAppEnvironment function resets environment.
    Use Get-Help Reset-PowerAppEnvironment -Examples for more detail.
    .PARAMETER EnvironmentName
    The environment name.
    .PARAMETER ResetRequestDefinition
    The ResetRequestDefinition object.
    .PARAMETER WaitUntilFinished
    If set to true, the function will not return until complete.
    .PARAMETER TimeoutInMinutes
    The timeout setting in minutes.
    .EXAMPLE
    Reset-PowerAppEnvironment -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -ResetRequestDefinition $resetRequest
    Resets environment 0fc02431-15fb-4563-a5ab-8211beb2a86f with $resetRequest object.
            $resetRequest = [pscustomobject]@{
            FriendlyName = "Friendly Name"
            DomainName = "url"
            Purpose = "purpose"
            BaseLanguageCode = 1
            Currency = [pscustomobject]@{
                Code = "USD"
                Name = "USD"
                Symbol = "$"
            }
            SecurityGroupId = "204162d5-59db-40c2-9788-2cda6b063f2b"
            Templates = @()
        }
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [object]$ResetRequestDefinition,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/reset?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

        $response = InvokeApi -Method POST -Body $ResetRequestDefinition -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        # Poll until the retore is completed
        If($WaitUntilFinished)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function Remove-PowerAppEnvironmentBackup
{
    <#
    .SYNOPSIS
    Remove environment bacup with backup Id.
    .DESCRIPTION
    The Remove-PowerAppEnvironmentBackup function removes environment bacup with backup Id.
    Use Get-Help Remove-PowerAppEnvironmentBackup -Examples for more detail.
    .PARAMETER EnvironmentName
    The environment name.
    .PARAMETER BackupId
    The environment backup Id.
    .EXAMPLE
    Remove-PowerAppEnvironmentBackup -EnvironmentName 0fc02431-15fb-4563-a5ab-8211beb2a86f -BackupId 942308ee-b1ba-433e-a08c-34b3d9ecaeef
    Remove environment 0fc02431-15fb-4563-a5ab-8211beb2a86f backup with backup Id 942308ee-b1ba-433e-a08c-34b3d9ecaeef.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$BackupId,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2019-05-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environment}/backups/{backupId}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName `
        | ReplaceMacro -Macro "{backupId}" -Value $BackupId;

        $response = InvokeApi -Method Delete -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-PowerAppTenantUrlPatterns
{
    <#
    .SYNOPSIS
    Retrieves a list of url patterns at tenant level.
    .DESCRIPTION
    Get-TenantUrlPatterns cmdlet gets url patterns for the logged in admin's tenant.
    Use Get-Help Get-TenantUrlPatterns -Examples for more detail.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-TenantUrlPatterns
    Retrieves all url patterns at the tenant level.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/urlPatterns"

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppPolicyUrlPatterns
{
    <#
    .SYNOPSIS
    Retrieves a list of url patterns for a DLP policy.
    .DESCRIPTION
    Get-PolicyUrlPatterns cmdlet gets policy url patterns for the logged in admin's policy.
    Use Get-Help Get-PolicyUrlPatterns -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PolicyUrlPatterns -TenantId $TenantId -PolicyName $PolicyName
    Retrieves url patterns for the given policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/urlPatterns" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppPolicyUrlPatterns
{
    <#
    .SYNOPSIS
    Creates new url patterns for a DLP policy by using NewUrlPatterns UrlPatternsDefinition object
    .DESCRIPTION
    New-PolicyUrlPatterns cmdlet creates new url patterns for the logged in admin's policy.
    Currently, this cmdlet is only enabled for polices at tenant level.
    Use Get-Help New-PolicyUrlPatterns -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER NewUrlPatterns
    Creates url patterns for a policy with NewUrlPatterns object.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    New-PolicyUrlPatterns -TenantId $TenantId -PolicyName $PolicyName -NewUrlPatterns $NewUrlPatterns;
    Creates url patterns for a policy with $TenantId, $PolicyName and $NewUrlPatterns object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$NewUrlPatterns,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        # get policy url patterns by policy name
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/urlPatterns" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        return InvokeApi -Method POST -Route $route -Body $NewUrlPatterns -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppTenantUrlPatterns
{
    <#
    .SYNOPSIS
    Creates new url patterns at the tenant by using NewUrlPatterns UrlPatternsDefinition object
    .DESCRIPTION
    New-TenantUrlPatterns cmdlet creates new url patterns for the logged in admin's tenant.
    Use Get-Help New-TenantUrlPatterns -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER NewUrlPatterns
    Creates url patterns at tenant level with NewUrlPatterns object.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    New-TenantUrlPatterns -NewUrlPatterns $NewUrlPatterns;
    Creates url patterns at tenant level with $NewUrlPatterns object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$NewUrlPatterns,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        # create api url patterns route for tenant
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/urlPatterns"

        return InvokeApi -Method POST -Route $route -Body $NewUrlPatterns -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-PowerAppTenantUrlPatterns
{
    <#
    .SYNOPSIS
    Deletes url patterns at tenant level.
    .DESCRIPTION
    Remove-TenantUrlPatterns cmdlet deletes the url patterns at tenant level.
    Use Get-Help Remove-TenantUrlPatterns -Examples for more detail.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Remove-DlpPolicy
    Deletes url patterns from tenant.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/urlPatterns"

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}


function Remove-PowerAppPolicyUrlPatterns
{
    <#
    .SYNOPSIS
    Deletes the specific url patterns by TenantId and PolicyName.
    .DESCRIPTION
    Remove-PolicyUrlPatterns cmdlet deletes url patterns from a specific policy.
    Use Get-Help Remove-PolicyUrlPatterns -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
    .PARAMETER PolicyName
    The PolicyName's identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Remove-PolicyUrlPatterns -TenantId $TenantId -PolicyName $PolicyName
    Deletes url patterns from policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/urlPatterns" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-PowerAppManagementApp
{
    <#
    .SYNOPSIS
    Get management application by Id.
    .DESCRIPTION
    Get-PowerAppManagementApp cmdlet returns the specified registered Azure Active Directory Web app/API application.
    Use Get-Help Get-PowerAppManagementApp -Examples for more detail.
    .PARAMETER ApplicationId
    The application's identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppManagementApp -ApplicationId $ApplicationId
    Get management application by ApplicationId.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$ApplicationId,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/adminApplications/{applicationId}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{applicationId}" -Value $ApplicationId;

        InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppManagementApps
{
    <#
    .SYNOPSIS
    Lists management applications.
    .DESCRIPTION
    Get-PowerAppManagementApps cmdlet returns a list of all registered Azure Active Directory Web app/API applications under the tenant.
    Use Get-Help Get-PowerAppManagementApps -Examples for more detail.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppManagementApps
    List management applications.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version={apiVersion}"

        InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppManagementApp
{
    <#
    .SYNOPSIS
    Register management application by Id.
    .DESCRIPTION
    New-PowerAppManagementApp cmdlet registers an Azure Active Directory Web app/API to be used with the Microsoft.Xrm.OnlineManagementAPI PowerShell module.
    Use Get-Help New-PowerAppManagementApp -Examples for more detail.
    .PARAMETER ApplicationId
    The application's identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    New-PowerAppManagementApp -ApplicationId $ApplicationId
    Register management application by ApplicationId.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$ApplicationId,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/adminApplications/{applicationId}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{applicationId}" -Value $ApplicationId;

        InvokeApi -Method PUT -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-PowerAppManagementApp
{
    <#
    .SYNOPSIS
    Remove management application by Id.
    .DESCRIPTION
    Remove-PowerAppManagementApp cmdlet removes the specified registered Azure Active Directory Web app/API application so it cannot be used with Microsoft.Xrm.OnlineManagementAPI PowerShell module.
    Use Get-Help Remove-PowerAppManagementApp -Examples for more detail.
    .PARAMETER ApplicationId
    The application's identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Remove-PowerAppManagementApp -ApplicationId $ApplicationId
    Remove management application by ApplicationId.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$ApplicationId,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/adminApplications/{applicationId}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{applicationId}" -Value $ApplicationId;

        InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-AdminPowerAppSharepointFormEnvironment
{
 <#
 .SYNOPSIS
 Gets the environment that Power Apps uses to save Sharepoint Form apps
 .DESCRIPTION
 Gets the environment Power Apps users to save Sharepoint form apps.
 Only tenant admin can call this route
 .PARAMETER ApiVersion
 The api version to call with. Default 2020-06-01
 .EXAMPLE
    Get-AdminPowerAppSharepointFormEnvironment
 #>
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environmentKeywords/~sharepointform?api-version={apiVersion}"

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-AdminPowerAppSharepointFormEnvironment
{
 <#
 .SYNOPSIS
 Remaps the environment that Power Apps uses to save Sharepoint Form apps
 .DESCRIPTION
 Modifies which environment Power Apps users to save Sharepoint form apps.
 Only tenant admin can call this route
 .PARAMETER EnvironmentName
 The Environment's identifier.
 .PARAMETER ApiVersion
 The api version to call with. Default 2020-06-01
 .EXAMPLE
    Set-AdminPowerAppSharepointFormEnvironment -EnvironmentName "14b9cc99-5714-45d6-bfe0-8377f4de9231"
 #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environmentKeywords/~sharepointform?api-version={apiVersion}"
        $requestBody = @{
            environmentName = $EnvironmentName
        }

        return InvokeApi -Method PUT -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Reset-AdminPowerAppSharepointFormEnvironment
{
 <#
 .SYNOPSIS
 Resets the environment that Power Apps uses to save Sharepoint Form apps back to the ~default environment
 .DESCRIPTION
 Gets the environment Power Apps users to save Sharepoint form apps back to the ~default environment.
 Only tenant admin can call this route
 .PARAMETER ApiVersion
 The api version to call with. Default 2020-06-01
 .EXAMPLE
    Reset-AdminPowerAppSharepointFormEnvironment
 #>
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environmentKeywords/~sharepointform?api-version={apiVersion}"

        return InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-AdminPowerAppTenantConsumedQuota
{
 <#
 .SYNOPSIS
 Gets the Admin PowerApps tenant consumed quota.
 .DESCRIPTION
 Gets the Admin PowerApps tenant consumed quota.
  .PARAMETER ApiVersion
 The api version to call with. Default 2020-06-01
 .EXAMPLE
    GetTenantConsumedQuota
 #>
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/scopes/admin/getConsumedQuota?api-version={apiVersion}"

        return InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppGenerateProtectionKey
{
    <#
    .SYNOPSIS
    Get PowerApp protection key.
    .DESCRIPTION
    Get-PowerAppGenerateProtectionKey cmdlet returns a new protection key.
    Use Get-Help Get-PowerAppGenerateProtectionKey -Examples for more detail.
    .PARAMETER LocationName
    The location name.
    .PARAMETER KeyName
    The key name.
    .PARAMETER KeyPassword
    The key password.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppGenerateProtectionKey -LocationName $LocationName -KeyName $KeyName -KeyPassword $KeyPassword
    Returns a new protection key.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$LocationName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$KeyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$KeyPassword,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/generateProtectionKey?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{location}" -Value $LocationName;

        $requestBody = @{
            KeyName = $KeyName
            KeyPassword = $KeyPassword
        }

        InvokeApi -Method POST -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppRetrieveTenantProtectionKey
{
    <#
    .SYNOPSIS
    Get PowerApp current protection key.
    .DESCRIPTION
    Get-PowerAppRetrieveTenantProtectionKey cmdlet returns the current tenant protection key.
    Use Get-Help Get-PowerAppRetrieveTenantProtectionKey -Examples for more detail.
    .PARAMETER LocationName
    The location name.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppRetrieveTenantProtectionKey -LocationName $LocationName -KeyName $KeyName -KeyPassword $KeyPassword
    Returns the current protection key.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$LocationName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/protectionKeys/~active?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{location}" -Value $LocationName;

        InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppRetrieveAvailableTenantProtectionKeys
{
    <#
    .SYNOPSIS
    Get PowerApp available keys.
    .DESCRIPTION
    Get-PowerAppRetrieveAvailableTenantProtectionKeys cmdlet returns the available protection keys for current tenant.
    Use Get-Help Get-PowerAppRetrieveAvailableTenantProtectionKeys -Examples for more detail.
    .PARAMETER LocationName
    The location name.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppRetrieveAvailableTenantProtectionKeys -LocationName $LocationName -KeyName $KeyName -KeyPassword $KeyPassword
    Returns available keys.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$LocationName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/protectionKeys?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{location}" -Value $LocationName;

        InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppImportProtectionKey
{
    <#
    .SYNOPSIS
    Import PowerApp protection key.
    .DESCRIPTION
    New-PowerAppImportProtectionKey cmdlet imports a new protection key.
    Use Get-Help New-PowerAppImportProtectionKey -Examples for more detail.
    .PARAMETER LocationName
    The location name.
    .PARAMETER KeyName
    The key name.
    .PARAMETER KeyType
    The key type.
    .PARAMETER KeyPassword
    The key password.
    .PARAMETER Key
    The key value.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    New-PowerAppImportProtectionKey -LocationName $LocationName -KeyName $KeyName -KeyType $KeyType -KeyPassword $KeyPassword -Key $Key
    Imports a new protection key.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$LocationName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$KeyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [int]$KeyType,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$KeyPassword,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [System.Byte[]] $Key,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/importProtectionKey?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{location}" -Value $LocationName;

        $requestBody = @{
            KeyName = $KeyName
            KeyType = $KeyType
            KeyPassword = $KeyPassword
            Key = $Key
        }

        $response = InvokeApi -Method POST -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Set-PowerAppProtectionStatus
{
    <#
     .SYNOPSIS
     Set environment protection status.
     .DESCRIPTION
     The Set-PowerAppProtectionStatus cmdlet updates environment protection status.
     Use Get-Help Set-PowerAppProtectionStatus -Examples for more detail.
     .PARAMETER EnvironmentName
     Updates a specific environment.
     .PARAMETER ProtectionKeyManagedBy
     Protection key managed by (Microsoft or Customer).
    .PARAMETER WaitUntilFinished
     Default is true.  If set to true, then the function will not return until operation completed.
    .PARAMETER TimeoutInMinutes
    The client timeout setting in minutes.
     .EXAMPLE
     Set-PowerAppProtectionStatus -EnvironmentName $EnvironmentName -ProtectionKeyManagedBy Microsoft
     Set environment protection managed by default Microsoft key.
     #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$ProtectionKeyManagedBy,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2019-10-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environmentName}/updateProtectionStatus`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentName}" -Value $EnvironmentName;

        $requestBody = @{
            managedBy = $ProtectionKeyManagedBy
        }

        $response = InvokeApiNoParseContent -Method POST -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($WaitUntilFinished -and $response.StatusCode -eq 202)
        {
            $response = WaitUntilFinished -Response $response -TimeoutInMinutes $TimeoutInMinutes -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }

        CreateHttpResponse($response)
    }
}

function Set-PowerAppTenantProtectionKey
{
    <#
    .SYNOPSIS
    Set PowerApp protection key.
    .DESCRIPTION
    Set-PowerAppTenantProtectionKey cmdlet sets an existing tenant protection key.
    Use Get-Help Set-PowerAppTenantProtectionKey -Examples for more detail.
    .PARAMETER LocationName
    The location name.
    .PARAMETER KeyName
    The key name.
    .PARAMETER WaitUntilFinished
     Default is false.  If set to true, then the function will not return until operation completed.
    .PARAMETER TimeoutInMinutes
    The client timeout setting in minutes.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppTenantProtectionKey -LocationName $LocationName -KeyName $KeyName
    Set protection key.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$LocationName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$KeyName,

        [Parameter(Mandatory = $false)]
        [bool]$WaitUntilFinished = $true,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutInMinutes = 10080, # Set max timeout to 1 week (60 min x 24 hour x 7 day)

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/protectionKeys/{keyName}/activate?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{location}" -Value $LocationName `
        | ReplaceMacro -Macro "{keyName}" -Value $KeyName;

        $response = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        If($WaitUntilFinished)
        {
            $statusUrl = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/lifecycleOperations/{operationId}?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{bapEndpoint}" -Value $global:currentSession.bapEndpoint `
            | ReplaceMacro -Macro "{operationId}" -Value $response.Id `
            | ReplaceMacro -Macro "{apiVersion}" -Value $ApiVersion

            $Headers = @{}
            $Headers['Operation-Location'] = $statusUrl

            $response = New-Object -TypeName PSObject `
                | Add-Member -PassThru -MemberType NoteProperty -Name StatusCode -Value 202 `
                | Add-Member -PassThru -MemberType NoteProperty -Name Headers -Value $Headers

            $response = WaitUntilFinished -Response $response -TimeoutInMinutes 60 -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }
        elseif ($response.state -eq "NotStarted")
        {
            return $response
        }

        return CreateHttpResponse($response)
    }
}

function Set-PowerAppLockAllEnvironments
{
    <#
    .SYNOPSIS
    Lock all environments.
    .DESCRIPTION
    Set-PowerAppLockAllEnvironments cmdlet locks all environments.
    Use Get-Help Set-PowerAppLockAllEnvironments -Examples for more detail.
    .PARAMETER LocationName
    The location name.
    .PARAMETER KeyName
    The key name.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppLockAllEnvironments -LocationName $LocationName
    Lock all environments.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$LocationName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$KeyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/locations/{location}/lockAllEnvironments?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{location}" -Value $LocationName;

        $requestBody = @{
            KeyName = $KeyName
        }

        $response = InvokeApi -Method POST -Route $route -Body $requestBody -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Set-PowerAppUnlockEnvironment
{
    <#
    .SYNOPSIS
    UnLock an environment.
    .DESCRIPTION
    Set-PowerAppUnlockEnvironment cmdlet unlocks an environment.
    Use Get-Help Set-PowerAppUnlockEnvironment -Examples for more detail.
    .PARAMETER $EnvironmentName
    The environment name.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppUnlockEnvironment -
    unlocks an environment.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$EnvironmentName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/Microsoft.BusinessAppPlatform/environments/{environmentname}/unlock?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{environmentname}" -Value $EnvironmentName;

        $response = InvokeApi -Method POST -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-PowerAppTenantIsolationPolicy
{
    <#
    .SYNOPSIS
    Get tenant isolation policy.
    .DESCRIPTION
    Get-PowerAppTenantIsolationPolicy cmdlet gets tenant isolation policy.
    Use Get-Help Get-PowerAppTenantIsolationPolicy -Examples for more detail.
    .PARAMETER TenantId
    The tenant Id.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppTenantIsolationPolicy -TenantId $TenantId
    Gets tenant isolation policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/tenantIsolationPolicy?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;

        InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-PowerAppTenantIsolationPolicy
{
    <#
    .SYNOPSIS
    Update tenant isolation policy.
    .DESCRIPTION
    Set-PowerAppTenantIsolationPolicy cmdlet updates tenant isolation policy.
    Use Get-Help Set-PowerAppTenantIsolationPolicy -Examples for more detail.
    .PARAMETER TenantId
    The tenant Id.
    .PARAMETER TenantIsolationPolicy
    The tenant isolation policy object.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppTenantIsolationPolicy -TenantId $TenantId
    Updates tenant isolation policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$TenantIsolationPolicy,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/tenantIsolationPolicy?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;

        InvokeApi -Method PUT -Route $route -Body $TenantIsolationPolicy -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppTenantIsolationOperationStatus
{
    <#
    .SYNOPSIS
    Get tenant isolation operation status.
    .DESCRIPTION
    Get-PowerAppTenantIsolationOperation cmdlet gets tenant isolation operation status.
    Use Get-Help Get-PowerAppTenantIsolationOperation -Examples for more detail.
    .PARAMETER TenantId
    The tenant Id.
    .PARAMETER OperationId
    The operation Id.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppTenantIsolationOperation -TenantId $TenantId -OperationId $OperationId
    Gets tenant isolation operation status.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$OperationId,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/tenantIsolationPolicy/operations/{operationId}?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{operationId}" -Value $OperationId;

        InvokeApiNoParseContent -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppDlpErrorSettings
{
    <#
    .SYNOPSIS
    Fetches the error settings for the tenant if present else returns null.
    .DESCRIPTION
    Get-PowerAppDlpErrorSettings cmdlet gets error settings for logged in tenant admin user.
    Use Get-Help Get-PowerAppDlpErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppDlpErrorSettings -TenantId $TenantId
    Fetches the error settings for the tenant if it exists else returns null
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppDlpErrorSettings
{
    <#
    .SYNOPSIS
    Creates the error settings for the tenant if its not created already.
    .DESCRIPTION
    New-PowerAppDlpErrorSettings cmdlet creates error settings for the tenant.
    Use Get-Help New-PowerAppDlpErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ErrorSettings
	The error settings object.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    New-PowerAppDlpErrorSettings -TenantId $TenantId -ErrorSettings $ErrorSettings
    Creates the error settings for the tenant based on the $ErrorSettings object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
		[Parameter(Mandatory = $true, ParameterSetName = "Name")]
		[Object]$ErrorSettings,
		[Parameter(Mandatory = $false, ParameterSetName = "Name")]
		[string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method POST -Route $route -Body $ErrorSettings -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-PowerAppDlpErrorSettings
{
    <#
    .SYNOPSIS
    Updates the error settings for the tenant if error settings are created for the tenant.
    .DESCRIPTION
    Set-PowerAppDlpErrorSettings cmdlet updates the error settings for the tenant admin if
	error settings exist for the tenant.
    Use Get-Help Set-PowerAppDlpErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ErrorSettings
	The error settings object.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppDlpErrorSettings -TenantId $TenantId -ErrorSettings $ErrorSettings
    Updates the error settings for the tenant.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
		[Parameter(Mandatory = $true, ParameterSetName = "Name")]
		[Object]$ErrorSettings,
		[Parameter(Mandatory = $false, ParameterSetName = "Name")]
		[string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method PATCH -Route $route -Body $ErrorSettings -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-PowerAppDlpErrorSettings
{
    <#
    .SYNOPSIS
    Deletes the error settings for the tenant if it exists.
    .DESCRIPTION
    Remove-PowerAppDlpErrorSettings cmdlet deletes the error settings for the tenant.
    Use Get-Help Remove-PowerAppDlpErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Remove-PowerAppDlpErrorSettings -TenantId $TenantId
    Deletes the error settings for the tenant if it exists.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-GovernanceErrorSettings
{
    <#
    .SYNOPSIS
    Fetches the error settings for the tenant if present else returns null.
    .DESCRIPTION
    Get-GovernanceErrorSettings cmdlet gets error settings for logged in tenant admin user.
    Use Get-Help Get-GovernanceErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Get-GovernanceErrorSettings -TenantId $TenantId
    Fetches the error settings for the tenant if it exists else returns null
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-GovernanceErrorSettings
{
    <#
    .SYNOPSIS
    Creates the error settings for the tenant if its not created already.
    .DESCRIPTION
    New-GovernanceErrorSettings cmdlet creates error settings for the tenant.
    Use Get-Help New-GovernanceErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ErrorSettings
	The error settings object.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    New-GovernanceErrorSettings -TenantId $TenantId -ErrorSettings $ErrorSettings
    Creates the error settings for the tenant based on the $ErrorSettings object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
		[Parameter(Mandatory = $true, ParameterSetName = "Name")]
		[Object]$ErrorSettings,
		[Parameter(Mandatory = $false, ParameterSetName = "Name")]
		[string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method POST -Route $route -Body $ErrorSettings -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-GovernanceErrorSettings
{
    <#
    .SYNOPSIS
    Updates the error settings for the tenant if error settings are created for the tenant.
    .DESCRIPTION
    Set-GovernanceErrorSettings cmdlet updates the error settings for the tenant admin if
	error settings exist for the tenant.
    Use Get-Help Set-GovernanceErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ErrorSettings
	The error settings object.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Set-GovernanceErrorSettings -TenantId $TenantId -ErrorSettings $ErrorSettings
    Updates the error settings for the tenant.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
		[Parameter(Mandatory = $true, ParameterSetName = "Name")]
		[Object]$ErrorSettings,
		[Parameter(Mandatory = $false, ParameterSetName = "Name")]
		[string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method PATCH -Route $route -Body $ErrorSettings -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-GovernanceErrorSettings
{
    <#
    .SYNOPSIS
    Deletes the error settings for the tenant if it exists.
    .DESCRIPTION
    Remove-GovernanceErrorSettings cmdlet deletes the error settings for the tenant.
    Use Get-Help Remove-GovernanceErrorSettings -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Remove-GovernanceErrorSettings -TenantId $TenantId
    Deletes the error settings for the tenant if it exists.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/errorSettings" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppDlpPolicyExemptResources
{
    <#
    .SYNOPSIS
    Retrieves exempt resources for a DLP policy.
    .DESCRIPTION
    Get-PowerAppDlpPolicyExemptResources cmdlet gets exempt resources object for the logged in admin's policy.
    Use Get-Help Get-PowerAppDlpPolicyExemptResources -Examples for more detail.
    .PARAMETER TenantId
    Get the specific exempt resources by using tenant id.
    .PARAMETER PolicyName
    Get the specific exempt resources by using policy name.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppDlpPolicyExemptResources -TenantId $TenantId -PolicyName $PolicyName
    Retrieves exempt resources for a DLP policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
         # get exempt resources by tenant id and policy name
            $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/exemptResources" `
            | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
            | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppDlpPolicyExemptResources
{
    <#
    .SYNOPSIS
    Creates new exempt resources for a DLP policy by using NewDlpPolicyExemptResources DlpPolicyExemptResourcesDefinition object
    .DESCRIPTION
    New-PowerAppDlpPolicyExemptResources cmdlet creates new cexempt resources for the logged in admin's DLP policy.
    Use Get-Help New-PowerAppDlpPolicyExemptResources -Examples for more detail.
    .PARAMETER TenantId
    Create new exempt resources by using tenant id.
    .PARAMETER PolicyName
    Create new exempt resources by using policy name.
    .PARAMETER NewDlpPolicyExemptResources
    Creates a DLP policy with NewDlpPolicyExemptResources object.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    New-PowerAppDlpPolicyExemptResources -TenantId $TenantId -PolicyName $PolicyName -NewDlpPolicyExemptResources $NewDlpPolicyExemptResources
    Creates a new policy with $TenantId, $PolicyName and $NewDlpPolicyExemptResources object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$PolicyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$NewDlpPolicyExemptResources,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/exemptResources" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        return InvokeApi -Method POST -Route $route -Body $NewDlpPolicyExemptResources -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-PowerAppDlpPolicyExemptResources
{
    <#
    .SYNOPSIS
    Deletes the specific DLP policy exempt resources by TenantId and PolicyName.
    .DESCRIPTION
    Remove-PowerAppDlpPolicyExemptResources cmdlet deletes exempt resources for a DLP policy.
    Use Get-Help Remove-PowerAppDlpPolicyExemptResources -Examples for more detail.
    .PARAMETER TenantId
    The policy with TenantId will be deleted.
    .PARAMETER PolicyName
    The policy with PolicyName will be deleted.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Remove-PowerAppDlpPolicyExemptResources -TenantId $TenantId -PolicyName $PolicyName
    Deletes exempt resources for a specific DLP policy.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name", ValueFromPipelineByPropertyName = $true)]
        [String]$PolicyName,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/exemptResources" `
            | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
            | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $response = InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Set-PowerAppDlpPolicyExemptResources
{
    <#
    .SYNOPSIS
    Updates exempt resources for a DLP policy by using UpdatedExemptResources DlpPolicyExemptResourcesDefinition object.
    .DESCRIPTION
    Set-PowerAppDlpPolicyExemptResources cmdlet updates details on the policy exempt resourcess.
    Use Get-Help Set-PowerAppDlpPolicyExemptResources -Examples for more detail.
    .PARAMETER TenantId
    Theexempt resources with TenantId will be updated.
    .PARAMETER PolicyName
    The exempt resources with PolicyName will be updated.
    .PARAMETER UpdatedExemptResources
    Policy exempt resources that will be updated.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppDlpPolicyExemptResources -UpdatedCExemptResources $UpdatedExemptResources
    Update the DLP policy exempt resources to $UpdatedExemptResources.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [String]$TenantId,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$PolicyName,

        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [object]$UpdatedExemptResources,

        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2016-11-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/policies/{policyName}/exemptResources" `
            | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
            | ReplaceMacro -Macro "{policyName}" -Value $PolicyName;

        $response = InvokeApi -Method PATCH -Route $route -Body $UpdatedExemptResources -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        CreateHttpResponse($response)
    }
}

function Get-PowerAppDlpConnectorBlockingPolicies
{
    <#
    .SYNOPSIS
    Fetches the connector blocking policies for the tenant if present else returns null.
    .DESCRIPTION
    Get-PowerAppDlpConnectorBlockingPolicies cmdlet gets connector blocking policies for logged in tenant admin user.
    Use Get-Help Get-PowerAppDlpConnectorBlockingPolicies -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppDlpConnectorBlockingPolicies -TenantId $TenantId
    Fetches the connector blocking policies for the tenant if it exists else returns null
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/connectorEnablementConfigurations" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Get-PowerAppDlpConnectorBlockingPolicy
{
    <#
    .SYNOPSIS
    Fetches the connector blocking policy for the tenant by Id.
    .DESCRIPTION
    Get-PowerAppDlpConnectorBlockingPolicy cmdlet gets connector blocking policy for the logged in user and the id mentioned.
    Use Get-Help Get-PowerAppDlpConnectorBlockingPolicy -Examples for more detail.
    .PARAMETER TenantId
    The TenantId identifier.
    .PARAMETER PolicyId
    The PolicyId identifier.
    .PARAMETER ApiVersion
    Specifies the Api version that is called.
    .EXAMPLE
    Get-PowerAppDlpConnectorBlockingPolicy -TenantId $tenantId -PolicyId $PolicyId
    Fetches the connector blocking policy for the logged in tenant admin recognized by id else returns null.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$PolicyId,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/connectorEnablementConfigurations/{id}" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{id}" -Value $PolicyId;
        return InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function New-PowerAppDlpConnectorBlockingPolicy
{
    <#
    .SYNOPSIS
    Creates the connector blocking policy for the tenant if its not created already.
    .DESCRIPTION
    New-PowerAppDlpConnectorBlockingPolicy cmdlet creates connector blocking policy for the tenant.
    Use Get-Help New-PowerAppDlpConnectorBlockingPolicy -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ConnectorBlockingDefinition
	The connector blocking policy definition.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    New-PowerAppDlpConnectorBlockingPolicy -TenantId $TenantId -ConnectorBlockingDefinition $ConnectorBlockingDefinition
    Creates the connector blocking policy for the tenant based on the $ConnectorBlockingDefinition object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
		[Parameter(Mandatory = $true, ParameterSetName = "Name")]
		[Object]$ConnectorBlockingDefinition,
		[Parameter(Mandatory = $false, ParameterSetName = "Name")]
		[string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $isDefinitionValid = IsConnectorBlockingPolicyDefinitionValid($ConnectorBlockingDefinition);

        if ($isDefinitionValid -eq $false)
        {
            return $null;
        }

        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/connectorEnablementConfigurations" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId;
        return InvokeApi -Method POST -Route $route -Body $ConnectorBlockingDefinition -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Set-PowerAppDlpConnectorBlockingPolicy
{
    <#
    .SYNOPSIS
    Updates the connector blocking policy for the tenant.
    .DESCRIPTION
    Set-PowerAppDlpConnectorBlockingPolicy cmdlet updates connector blocking policy for the tenant.
    Use Get-Help Set-PowerAppDlpConnectorBlockingPolicy -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
	.PARAMETER ConnectorBlockingDefinition
	The connector blocking policy definition.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Set-PowerAppDlpConnectorBlockingPolicy -TenantId $TenantId -PolicyId $PolicyId -ConnectorBlockingDefinition $ConnectorBlockingDefinition
    Updates the connector blocking policy for the tenant based on the $ConnectorBlockingDefinition object.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$PolicyId,
		[Parameter(Mandatory = $true, ParameterSetName = "Name")]
		[Object]$ConnectorBlockingDefinition,
		[Parameter(Mandatory = $false, ParameterSetName = "Name")]
		[string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $isDefinitionValid = IsConnectorBlockingPolicyDefinitionValid($ConnectorBlockingDefinition);

        if ($isDefinitionValid -eq $false)
        {
            return $null;
        }
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/connectorEnablementConfigurations/{policyId}" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{policyId}" -Value $PolicyId;
        return InvokeApi -Method PATCH -Route $route -Body $ConnectorBlockingDefinition -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

function Remove-PowerAppDlpConnectorBlockingPolicy
{
    <#
    .SYNOPSIS
    Deletes the connector blocking policy for the tenant if it exists.
    .DESCRIPTION
    Remove-PowerAppDlpConnectorBlockingPolicy cmdlet deletes the connector blocking policy for the tenant.
    Use Get-Help Remove-PowerAppDlpConnectorBlockingPolicy -Examples for more detail.
    .PARAMETER TenantId
    The TenantId's identifier.
    .PARAMETER policyId
    The PolicyId's identifier.
	.PARAMETER ApiVersion
	Specifies the Api version that is called.
    .EXAMPLE
    Remove-PowerAppDlpConnectorBlockingPolicy -TenantId $TenantId -PolicyId $PolicyId
    Deletes the connector blocking policy for the tenant if it exists.
    #>
    [CmdletBinding(DefaultParameterSetName="Name")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$TenantId,
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]$PolicyId,
        [Parameter(Mandatory = $false, ParameterSetName = "Name")]
        [string]$ApiVersion = "2020-06-01"
    )
    process
    {
        $route = "https://{bapEndpoint}/providers/PowerPlatform.Governance/v1/tenants/{tenantId}/connectorEnablementConfigurations/{policyId}" `
        | ReplaceMacro -Macro "{tenantId}" -Value $TenantId `
        | ReplaceMacro -Macro "{policyId}" -Value $PolicyId;
        return InvokeApi -Method DELETE -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
}

#internal, helper function
function Get-StatusMessage {
    param (
        [int]$StatusCode
    )

    switch ($StatusCode) {
        200 { return "Succeeded" }
        202 { return "InProgress" }
        default { return "Failed" }
    }
}

#internal, helper function
function CreatePreflightResponse
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ResponseObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name HttpStatusCode -Value $ResponseObject.StatusCode `
        | Add-Member -PassThru -MemberType NoteProperty -Name ErrorCode -Value $ResponseObject.error.code `
        | Add-Member -PassThru -MemberType NoteProperty -Name ErrorMessage -Value $ResponseObject.error.message `
}

#internal, helper function
function CreateMigrationOperation
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ResponseObject
    )

    if ($result -and $result.Headers['Location']){
        $nextLink = $result.Headers['Location'] | Select-Object -First 1
    } else {
        $nextLink = $null
    }

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name Status -Value $ResponseObject.StatusCode `
        | Add-Member -PassThru -MemberType NoteProperty -Name NextLink -Value $nextLink `
        | Add-Member -PassThru -MemberType NoteProperty -Name Error -Value $ResponseObject.error.message `
}

#internal, helper function
function Get-FilteredEnvironments
{
    param
    (
        [Parameter(Mandatory = $false)]
        [object]$Filter,

        [Parameter(Mandatory = $false)]
        [object]$CreatedBy,

        [Parameter(Mandatory = $false)]
        [object]$EnvironmentResult,

        [Parameter(Mandatory = $false)]
        [bool]$ReturnCdsDatabaseType = $false,

        [Parameter(Mandatory = $false)]
        [bool]$GetGenerativeAiSettings,

        [Parameter(Mandatory = $false)]
        [bool]$GetMakerAnalyticsSettings
    )

    $patternOwner = BuildFilterPattern -Filter $CreatedBy
    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($env in $EnvironmentResult.Value)
    {
        if ($patternOwner.IsMatch($env.properties.createdBy.displayName) -or
            $patternOwner.IsMatch($env.properties.createdBy.email) -or
            $patternOwner.IsMatch($env.properties.createdBy.id) -or
            $patternOwner.IsMatch($env.properties.createdBy.userPrincipalName))
        {
            if ($patternFilter.IsMatch($env.name) -or
                $patternFilter.IsMatch($env.properties.displayName))
            {
                CreateEnvironmentObject -EnvObject $env -ReturnCdsDatabaseType $ReturnCdsDatabaseType -GetGenerativeAiSettings $GetGenerativeAiSettings -GetMakerAnalyticsSettings $GetMakerAnalyticsSettings
            }
        }
    }
}

#internal, helper function
function Get-FilteredApiPolicies
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ApiPolicyResult,

        [Parameter(Mandatory = $false)]
        [string]$CreatedBy,

        [Parameter(Mandatory = $false)]
        [string]$PolicyName,

        [Parameter(Mandatory = $false)]
        [string]$Filter
    )

    $patternPolicyName = BuildFilterPattern -Filter $PolicyName
    $patternOwner = BuildFilterPattern -Filter $CreatedBy
    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($pol in $ApiPolicyResult.Value)
    {
        if ($patternPolicyName.IsMatch($pol.name))
        {
            if ($patternOwner.IsMatch($pol.properties.createdBy.displayName) -or
                $patternOwner.IsMatch($pol.properties.createdBy.email) -or
                $patternOwner.IsMatch($pol.properties.createdBy.id) -or
                $patternOwner.IsMatch($pol.properties.createdBy.userPrincipalName))
            {
                if ($patternFilter.IsMatch($pol.name) -or
                    $patternFilter.IsMatch($pol.properties.displayName))
                {
                    CreateApiPolicyObject -PolicyObject $pol
                }
            }
        }
    }
}

#internal, helper function
function Get-CdsOneDatabase(
    [string]$ApiVersion = "2016-11-01",
    [string]$EnvironmentName = "2016-11-01"

)
{
    $route = "https://{cdsOneEndpoint}/providers/Microsoft.CommonDataModel/namespaces?api-version={apiVersion}&`$filter=environment%20eq%20%27{environment}%27" `
    | ReplaceMacro -Macro "{environment}" -Value $EnvironmentName;

    $databaseResult = InvokeApi -Method GET -Route $route -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    CreateCdsOneDatabasObject -DatabaseObject $databaseResult.Value
}

#internal, helper function
function Get-FilteredApps
{
     param
    (
        [Parameter(Mandatory = $false)]
        [object]$Filter,

        [Parameter(Mandatory = $false)]
        [object]$Owner,

        [Parameter(Mandatory = $false)]
        [object]$AppResult
    )

    $patternOwner = BuildFilterPattern -Filter $Owner
    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($app in $AppResult.Value)
    {
        if ($patternOwner.IsMatch($app.properties.owner.displayName) -or
            $patternOwner.IsMatch($app.properties.owner.email) -or
            $patternOwner.IsMatch($app.properties.owner.id) -or
            $patternOwner.IsMatch($app.properties.owner.userPrincipalName))
        {
            if ($patternFilter.IsMatch($app.name) -or
                $patternFilter.IsMatch($app.properties.displayName))
            {
                CreateAppObject -AppObj $app
            }
        }
    }
}

#internal, helper function
function Get-FilteredConnections
{
     param
    (
        [Parameter(Mandatory = $false)]
        [object]$Filter,

        [Parameter(Mandatory = $false)]
        [object]$CreatedBy,

        [Parameter(Mandatory = $false)]
        [object]$ConnectionResult
    )

    $patternCreatedBy = BuildFilterPattern -Filter $CreatedBy
    $patternFilter = BuildFilterPattern -Filter $Filter

    foreach ($connection in $ConnectionResult.Value)
    {
        if ($patternCreatedBy.IsMatch($connection.properties.createdBy.displayName) -or
            $patternCreatedBy.IsMatch($connection.properties.createdBy.email) -or
            $patternCreatedBy.IsMatch($connection.properties.createdBy.id) -or
            $patternCreatedBy.IsMatch($connection.properties.createdBy.userPrincipalName))
        {
            if ($patternFilter.IsMatch($connection.name) -or
                $patternFilter.IsMatch($connection.properties.displayName))
            {
                CreateConnectionObject -ConnectionObj $connection
            }
        }
    }
}

#internal, helper function
function Get-FilteredFlows
{
    param
    (
        [Parameter(Mandatory = $false)]
        [object]$Filter,

        [Parameter(Mandatory = $false)]
        [object]$CreatedBy,

        [Parameter(Mandatory = $false)]
        [object]$FlowResult
    )

    if ($FlowResult.Value -ne $null -and $FlowResult.Value.Count -gt 0)
    {
        if (-not [string]::IsNullOrWhiteSpace($CreatedBy))
        {
            $pattern = BuildFilterPattern -Filter $CreatedBy

            foreach ($flow in $FlowResult.Value)
            {
                if ($flow -ne $null -and
                    ($pattern.IsMatch($flow.properties.creator.objectId) -or
                    $pattern.IsMatch($flow.properties.creator.userId)))
                {
                    CreateFlowObject -FlowObj $flow
                }
            }
        }
        else
        {
            $pattern = BuildFilterPattern -Filter $Filter

            foreach ($flow in $FlowResult.Value)
            {
                if ($flow -ne $null -and
                    ($pattern.IsMatch($flow.name) -or
                    $pattern.IsMatch($flow.properties.displayName)))
                {
                    CreateFlowObject -FlowObj $flow
                }
            }
        }
    }
}


#internal, helper function
function CreateHttpResponse
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ResponseObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name Code -Value $ResponseObject.StatusCode `
        | Add-Member -PassThru -MemberType NoteProperty -Name Description -Value $ResponseObject.StatusDescription `
        | Add-Member -PassThru -MemberType NoteProperty -Name Headers -Value $ResponseObject.Headers `
        | Add-Member -PassThru -MemberType NoteProperty -Name Error -Value $ResponseObject.error `
        | Add-Member -PassThru -MemberType NoteProperty -Name Errors -Value $ResponseObject.errors `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $ResponseObject;
}

#internal, helper function
function CreateEnvironmentObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$EnvObject,

        [Parameter(Mandatory = $false)]
        [bool]$ReturnCdsDatabaseType,

        [Parameter(Mandatory = $false)]
        [bool]$GetGenerativeAiSettings,

        [Parameter(Mandatory = $false)]
        [bool]$GetMakerAnalyticsSettings
    )

    If($ReturnCdsDatabaseType)
    {
        $cdsDatabaseType = "None"

        # this property will be set if the environment has linked CDS 2.0 database
        $LinkedCdsTwoInstanceType = $EnvObject.properties.linkedEnvironmentMetadata.type;

        if ($LinkedCdsTwoInstanceType -ne $null)
        {
            if($LinkedCdsTwoInstanceType -eq "Dynamics365Instance")
            {
                $cdsDatabaseType = "Common Data Service for Apps"
            }
            else
            {
                #unfortunately there is no other way to determine if an environment has a database other than making a separate REST API call
                $cdsOneDatabase = Get-CdsOneDatabase -ApiVersion $ApiVersion -EnvironmentName $EnvObject.name

                if ($cdsOneDatabase.EnvironmentName -eq $EnvObject.name)
                {
                    $cdsDatabaseType = "Common Data Service (Previous Version)"
                }
            }
        }
    }
    else {
        $cdsDatabaseType = "Unknown"
    }

    $retentionPeriod = [System.Xml.XmlConvert]::ToTimeSpan($EnvObject.properties.retentionDetails.retentionPeriod).Days

    $newEnvObject = New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $EnvObject.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $EnvObject.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name Description -Value $EnvObject.properties.description `
        | Add-Member -PassThru -MemberType NoteProperty -Name IsDefault -Value $EnvObject.properties.isDefault `
        | Add-Member -PassThru -MemberType NoteProperty -Name Location -Value $EnvObject.location `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $EnvObject.properties.createdTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedBy -value $EnvObject.properties.createdBy `
        | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedTime -Value $EnvObject.properties.lastModifiedTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedBy -value $EnvObject.properties.lastModifiedBy.userPrincipalName `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreationType -value $EnvObject.properties.creationType `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentType -value $EnvObject.properties.environmentSku `
        | Add-Member -PassThru -MemberType NoteProperty -Name CommonDataServiceDatabaseProvisioningState -Value $EnvObject.properties.provisioningState `
        | Add-Member -PassThru -MemberType NoteProperty -Name CommonDataServiceDatabaseType -Value $cdsDatabaseType `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $EnvObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name InternalCds -value $cdsOneDatabase;

    if ($EnvObject.properties.linkedEnvironmentMetadata -ne $null)
    {
        Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name OrganizationId -Value $EnvObject.properties.linkedEnvironmentMetadata.resourceId;

        if ($EnvObject.properties.linkedEnvironmentMetadata.securityGroupId -ne $null)
        {
            Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name SecurityGroupId -Value $EnvObject.properties.linkedEnvironmentMetadata.securityGroupId;
        }
    }

    if ($null -ne $EnvObject.properties.capacity)
    {
        Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name Capacity -Value $EnvObject.properties.capacity
    }

    if (0 -lt $retentionPeriod)
    {
        Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name RetentionPeriod -Value $retentionPeriod
    }

    if ($GetGenerativeAiSettings)
    {
        if ($null -ne $EnvObject.properties.copilotPolicies)
        {
            Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name CrossGeoDataMovement -Value $EnvObject.properties.copilotPolicies.crossGeoCopilotDataMovementEnabled
        }
        else
        {
            Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name CrossGeoDataMovement -Value 'N/A (data is not processed out of geo)'
        }

        Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name BingSearch -Value $EnvObject.properties.bingChatEnabled
        Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name Microsoft365Services -Value $EnvObject.properties.m365Enabled
    }

    if ($GetMakerAnalyticsSettings)
    {
        if ($null -ne $EnvObject.properties.enableAnalyticsForMakers)
        {
            Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name EnableAnalyticsForMakers -Value $EnvObject.properties.enableAnalyticsForMakers
        }
        if($null -ne $EnvObject.properties.enableEUPIForMakerAnalytics)
        {
            Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name EnableEUPIForMakerAnalytics -Value $EnvObject.properties.enableEUPIForMakerAnalytics
        }
        else
        {
            Add-Member -InputObject $newEnvObject -MemberType NoteProperty -Name EnableEUPIForMakerAnalytics -Value 'null (null is treated as false)'
        }
    }

    return $newEnvObject;
}

#internal, helper function
function CreateEnvironmentLocationObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$EnvironmentLocationObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name LocationName -Value $EnvironmentLocationObject.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name LocationDisplayName -Value $EnvironmentLocationObject.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name AzureRegions -Value $EnvironmentLocationObject.properties.azureRegions `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $EnvironmentLocationObject;
}

#internal, helper function
function CreateCurrencyObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$CurrencyObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name CurrencyName -Value $CurrencyObject.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name CurrencyCode -Value $CurrencyObject.properties.code `
        | Add-Member -PassThru -MemberType NoteProperty -Name IsTenantDefaultCurrency -Value $CurrencyObject.properties.isTenantDefault `
        | Add-Member -PassThru -MemberType NoteProperty -Name CurrencySymbol -Value $CurrencyObject.properties.symbol `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $CurrencyObject;
}

#internal, helper function
function CreateLanguageObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$LanguageObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name LanguageName -Value $LanguageObject.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name LanguageDisplayName -Value $LanguageObject.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name IsTenantDefaultLanguage -Value $LanguageObject.properties.isTenantDefault `
        | Add-Member -PassThru -MemberType NoteProperty -Name LanguageLocalizedDisplayName -Value $LanguageObject.properties.localizedName `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $LanguageObject;
}

#internal, helper function
function CreateTemplateObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$TemplateObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name TemplateName -Value $TemplateObject.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name TemplateLocation -Value $TemplateObject.location `
        | Add-Member -PassThru -MemberType NoteProperty -Name TemplateDisplayName -Value $TemplateObject.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name IsDisabled -Value $TemplateObject.properties.isDisabled
}

#internal, helper function
function CreateCdsOneDatabasObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$DatabaseObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name DatabaseId -Value $DatabaseObject.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name DatabaseName -Value $DatabaseObject.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $DatabaseObject.properties.environmentId `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $DatabaseObject.properties.createdDateTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $DatabaseObject.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name ProvisioningState -Value $DatabaseObject.properties.provisioningState `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $DatabaseObject;
}

#internal, helper function
function CreateAppObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$AppObj
    )

    if ($AppObj.name -ne $null)
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name AppName -Value $AppObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $AppObj.properties.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $AppObj.properties.createdTime `
            | Add-Member -PassThru -MemberType NoteProperty -Name Owner -Value $AppObj.properties.owner `
            | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedTime -Value $AppObj.properties.lastModifiedTime `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $AppObj.properties.environment.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name UnpublishedAppDefinition -Value $AppObj.properties.unpublishedAppDefinition `
            | Add-Member -PassThru -MemberType NoteProperty -Name IsFeaturedApp -Value $AppObj.properties.isFeaturedApp `
            | Add-Member -PassThru -MemberType NoteProperty -Name IsHeroApp -Value $AppObj.properties.isHeroApp `
            | Add-Member -PassThru -MemberType NoteProperty -Name BypassConsent -Value $AppObj.properties.bypassConsent `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $AppObj;
            #bypassConsent
    }
    else
    {
        return $AppObj
    }
}

#internal, helper function
function CreateFlowObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$FlowObj
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name FlowName -Value $FlowObj.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name Enabled -Value ($FlowObj.properties.state -eq 'Started') `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $FlowObj.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name UserType -Value $FlowObj.properties.userType `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $FlowObj.properties.createdTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedBy -Value $FlowObj.properties.creator `
        | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedTime -Value $FlowObj.properties.lastModifiedTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $FlowObj.properties.environment.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name WorkflowEntityId -Value $FlowObj.properties.workflowEntityId `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $FlowObj;
}

#internal, helper function
function CreateAppRoleAssignmentObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$AppRoleAssignmentObj
    )

    If($AppRoleAssignmentObj.properties.principal.type -eq "Tenant")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $AppRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $AppRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $AppRoleAssignmentObj.properties.principal.tenantId `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $AppRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $AppRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name AppName -Value ((($AppRoleAssignmentObj.properties.scope -split "/apps/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($AppRoleAssignmentObj.properties.scope -split "/environments/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $AppRoleAssignmentObj;
    }
    elseif($AppRoleAssignmentObj.properties.principal.type -eq "User")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $AppRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $AppRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $AppRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $AppRoleAssignmentObj.properties.principal.email `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $AppRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $AppRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $AppRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name AppName -Value ((($AppRoleAssignmentObj.properties.scope -split "/apps/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($AppRoleAssignmentObj.properties.scope -split "/environments/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $AppRoleAssignmentObj;
    }
    elseif($AppRoleAssignmentObj.properties.principal.type -eq "Group")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $AppRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $AppRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $AppRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $AppRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $AppRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $AppRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name AppName -Value ((($AppRoleAssignmentObj.properties.scope -split "/apps/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($AppRoleAssignmentObj.properties.scope -split "/environments/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $AppRoleAssignmentObj;
    }
    else {
        return $null
    }
}

#internal, helper function
function CreateFlowRoleAssignmentObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$FlowRoleAssignmentObj
    )

    if($FlowRoleAssignmentObj.properties.principal.type -eq "User")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $FlowRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $FlowRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $FlowRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $FlowRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $FlowRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name FlowName -Value ((($FlowRoleAssignmentObj.id -split "/flows/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($FlowRoleAssignmentObj.id -split "/environments/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $FlowRoleAssignmentObj;
    }
    elseif($FlowRoleAssignmentObj.properties.principal.type -eq "Group")
    {
        return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $FlowRoleAssignmentObj.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $FlowRoleAssignmentObj.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $FlowRoleAssignmentObj.properties.principal.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $FlowRoleAssignmentObj.properties.principal.type `
        | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $FlowRoleAssignmentObj.properties.roleName `
        | Add-Member -PassThru -MemberType NoteProperty -Name FlowName -Value ((($FlowRoleAssignmentObj.id -split "/flows/")[1]) -split "/")[0] `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($FlowRoleAssignmentObj.id -split "/environments/")[1]) -split "/")[0] `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $FlowRoleAssignmentObj;
    }
    else {
        return $null
    }
}

#internal, helper function
function CreateEnvRoleAssignmentObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$EnvRoleAssignmentObj,

        [Parameter(Mandatory = $false)]
        [object]$EnvObj
    )

    If($EnvRoleAssignmentObj.properties.principal.type -eq "Tenant")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $EnvRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $EnvRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $EnvRoleAssignmentObj.properties.principal.tenantId `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $EnvRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $EnvRoleAssignmentObj.properties.roleDefinition.name`
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($EnvRoleAssignmentObj.properties.scope -split "/environments/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentObject -Value $EnvObj `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $EnvRoleAssignmentObj;
    }
    elseif($EnvRoleAssignmentObj.properties.principal.type -eq "User")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $EnvRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $EnvRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $EnvRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $EnvRoleAssignmentObj.properties.principal.email `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $EnvRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $EnvRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $EnvRoleAssignmentObj.properties.roleDefinition.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($EnvRoleAssignmentObj.properties.scope -split "/environments/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentObject -Value $EnvObj `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $EnvRoleAssignmentObj;
    }
    elseif($EnvRoleAssignmentObj.properties.principal.type -eq "Group")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $EnvRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $EnvRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $EnvRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $EnvRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $EnvRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $EnvRoleAssignmentObj.properties.roleDefinition.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value ((($EnvRoleAssignmentObj.properties.scope -split "/environments/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentObject -Value $EnvObj `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $EnvRoleAssignmentObj;
    }
    else {
        return $null
    }
}

#internal, helper function
function CreateConnectionObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ConnectionObj
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name ConnectionName -Value $ConnectionObj.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name ConnectionId -Value $ConnectionObj.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name FullConnectorName -Value $ConnectionObj.properties.apiId `
        | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value ((($ConnectionObj.properties.apiId -split "/apis/")[1]) -split "/")[0] `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $ConnectionObj.properties.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $ConnectionObj.properties.createdTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name CreatedBy -Value $ConnectionObj.properties.createdBy `
        | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedTime -Value $ConnectionObj.properties.lastModifiedTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $ConnectionObj.properties.environment.name `
        | Add-Member -PassThru -MemberType NoteProperty -Name Statuses -Value $ConnectionObj.properties.statuses `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $ConnectionObj;
}

#internal, helper function
function CreateConnectionRoleAssignmentObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ConnectionRoleAssignmentObj,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName
    )

    If($ConnectionRoleAssignmentObj.properties.principal.type -eq "Tenant")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $ConnectionRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $ConnectionRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $ConnectionRoleAssignmentObj.properties.principal.tenantId `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $ConnectionRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $ConnectionRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectionName -Value ((($ConnectionRoleAssignmentObj.id -split "/connections/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value ((($ConnectionRoleAssignmentObj.id -split "/apis/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $EnvironmentName `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $ConnectionRoleAssignmentObj;
    }
    elseif($ConnectionRoleAssignmentObj.properties.principal.type -eq "User")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $ConnectionRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $ConnectionRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $ConnectionRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $ConnectionRoleAssignmentObj.properties.principal.email `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $ConnectionRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $ConnectionRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $ConnectionRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectionName -Value ((($ConnectionRoleAssignmentObj.id -split "/connections/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value ((($ConnectionRoleAssignmentObj.id -split "/apis/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $EnvironmentName `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $ConnectionRoleAssignmentObj;
    }
    elseif($ConnectionRoleAssignmentObj.properties.principal.type -eq "Group")
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleId -Value $ConnectionRoleAssignmentObj.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleName -Value $ConnectionRoleAssignmentObj.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalDisplayName -Value $ConnectionRoleAssignmentObj.properties.principal.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalEmail -Value $null `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalObjectId -Value $ConnectionRoleAssignmentObj.properties.principal.id `
            | Add-Member -PassThru -MemberType NoteProperty -Name PrincipalType -Value $ConnectionRoleAssignmentObj.properties.principal.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name RoleType -Value $ConnectionRoleAssignmentObj.properties.roleName `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectionName -Value ((($ConnectionRoleAssignmentObj.id -split "/permission/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name ConnectorName -Value ((($ConnectionRoleAssignmentObj.id -split "/apis/")[1]) -split "/")[0] `
            | Add-Member -PassThru -MemberType NoteProperty -Name EnvironmentName -Value $EnvironmentName `
            | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $ConnectionRoleAssignmentObj;
    }
    else {
        return $null
    }
}

#internal, helper function
function CreateFlowUserDetailsObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$FlowUserObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name ConsentBusinessAppPlatformTime -Value $FlowUserObject.consentBusinessAppPlatformTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name ConsentTime -Value $FlowUserObject.consentTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name IsDisallowedForInternalPlans -Value $FlowUserObject.isDisallowedForInternalPlans `
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectId -Value $FlowUserObject.objectId `
        | Add-Member -PassThru -MemberType NoteProperty -Name Puid -Value $FlowUserObject.puid `
        | Add-Member -PassThru -MemberType NoteProperty -Name ServiceSettingsSelectionTime -Value $FlowUserObject.serviceSettingsSelectionTime `
        | Add-Member -PassThru -MemberType NoteProperty -Name TenantId -Value $FlowUserObject.tenantId;
}

#internal, helper function
function CreateConnectorActionObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ActionObject
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name Id -Value $ActionObject.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name Type -Value $ActionObject.type `
        | Add-Member -PassThru -MemberType NoteProperty -Name Properties -Value $ActionObject.properties;
}

#internal, helper method
function CreateApiPolicyObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$PolicyObject
    )

    if ($PolicyObject.name -ne $null)
    {
        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name PolicyName -Value $PolicyObject.name `
            | Add-Member -PassThru -MemberType NoteProperty -Name Type -Value $PolicyObject.type `
            | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $PolicyObject.properties.displayName `
            | Add-Member -PassThru -MemberType NoteProperty -Name CreatedTime -Value $PolicyObject.properties.createdTime `
            | Add-Member -PassThru -MemberType NoteProperty -Name CreatedBy -Value $PolicyObject.properties.createdBy `
            | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedTime -Value $PolicyObject.properties.lastModifiedTime `
            | Add-Member -PassThru -MemberType NoteProperty -Name LastModifiedBy -Value $PolicyObject.properties.lastModifiedBy `
            | Add-Member -PassThru -MemberType NoteProperty -Name Constraints -Value $PolicyObject.properties.definition.constraints `
            | Add-Member -PassThru -MemberType NoteProperty -Name BusinessDataGroup -Value $PolicyObject.properties.definition.apiGroups.hbi.apis`
            | Add-Member -PassThru -MemberType NoteProperty -Name NonBusinessDataGroup -Value $PolicyObject.properties.definition.apiGroups.lbi.apis`
            | Add-Member -PassThru -MemberType NoteProperty -Name BlockedGroup -Value $PolicyObject.properties.definition.apiGroups.blocked.apis`
            | Add-Member -PassThru -MemberType NoteProperty -Name FilterType -Value $PolicyObject.properties.definition.constraints.environmentFilter1.parameters.filterType `
            | Add-Member -PassThru -MemberType NoteProperty -Name Environments -Value $PolicyObject.properties.definition.constraints.environmentFilter1.parameters.environments;
    }
    else
    {
        return $PolicyObject
    }
}

#internal, helper method
function AcquireLeaseAndPutApp(
    [CmdletBinding()]

    [string]$AppName,
    [string]$ApiVersion,
    [Object]$PowerApp,
    [Boolean]$ForceLease
)
{
    if ($ApiVersion -eq $null -or $ApiVersion -eq "")
    {
        Write-Error "Api Version must be set."
        throw
    }

    $apiVersionsBeforePublishSave = @("2016-11-01", "2017-02-01", "2017-04-01")
    foreach ($apiVersionPrefix in $apiVersionsBeforePublishSave)
    {
        $doesNotNeedPublish = $ApiVersion -Match $apiVersionPrefix
        if ($doesNotNeedPublish)
        {
            Write-Warning "Older API version, please use 2017-05-01 or newer."
            break;
        }
    }

    if ($ForceLease)
    {
        $forceLeaseFlag = "true"
    }
    else
    {
        $forceLeaseFlag = "false"
    }

    $powerAppBaseUri = "https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/apps/{appName}" `
        | ReplaceMacro -Macro "{powerAppsEndpoint}" -Value $Global:currentSession.powerAppsEndpoint `
        | ReplaceMacro -Macro "{appName}" -Value $AppName;

    $acquireLeaseUri = "{powerAppBaseUri}/acquireLease`?api-version={apiVersion}&forceLeaseAcquisition={forceLeaseFlag}" `
        | ReplaceMacro -Macro "{powerAppBaseUri}" -Value $powerAppBaseUri `
        | ReplaceMacro -Macro "{forceLeaseFlag}" -Value $forceLeaseFlag;

    $releaseLeaseUri = "{powerAppBaseUri}/releaseLease`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{powerAppBaseUri}" -Value $powerAppBaseUri;

    $putPowerAppUri = "{powerAppBaseUri}/`?api-version={apiVersion}" `
        | ReplaceMacro -Macro "{powerAppBaseUri}" -Value $powerAppBaseUri;

    $leaseResponse = InvokeApi -Route $acquireLeaseUri -Method Post -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

    if ($doesNotNeedPublish)
    {
        $response = InvokeApi -Route $putPowerAppUri -Method Put -Body $PowerApp -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }
    else
    {
        $powerApp.Properties.LifeCycleId = "Draft"
        $response = InvokeApi -Route $putPowerAppUri -Method Put -Body $PowerApp -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

        $publishPowerAppUri = "{powerAppBaseUri}/publish`?api-version={apiVersion}" `
            | ReplaceMacro -Macro "{powerAppBaseUri}" -Value $powerAppBaseUri `
            | ReplaceMacro -Macro "{apiVersion}" -Value $ApiVersion;

        $publishResponse = InvokeApi -Route $publishPowerAppUri -Method Post -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    }

    $response = InvokeApi -Route $releaseLeaseUri -Method Post -Body $leaseResponse -ThrowOnFailure -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    CreateHttpResponse($response)
}

#internal, helper method
function WaitUntilFinished(
    [CmdletBinding()]
    [Parameter(Mandatory = $true)]
    [object]$Response,

    [Parameter(Mandatory = $true)]
    [int]$TimeoutInMinutes,

    [Parameter(Mandatory = $true)]
    [string]$ApiVersion
)
{
    $response = $Response
    $statusUrl = $response.Headers['Operation-Location']
    if ([string]::IsNullOrEmpty($statusUrl))
    {
        $statusUrl = $response.Headers['Location']
    }

    if ($Response.StatusCode -eq 202)
    {
        $currentTime = Get-Date -format HH:mm:ss
        $nextTime = Get-Date -format HH:mm:ss
        $TimeDiff = New-TimeSpan $currentTime $nextTime
        $state = "NotStarted"

        #Wait until the operation completed or the service timeout
        while((-not [string]::IsNullOrEmpty($statusUrl)) -and ($state -ne "Succeeded") -and ($state -ne "Failed") -and ($TimeDiff.TotalMinutes -lt $TimeoutInMinutes))
        {
            Start-Sleep -s 5
            $response = InvokeApiNoParseContent -Route $statusUrl -Method GET -ApiVersion $ApiVersion -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
            $nextTime = Get-Date -format HH:mm:ss
            $TimeDiff = New-TimeSpan $currentTime $nextTime
            $content = ConvertFrom-Json $response.Content
            $state = $content.state.id.ToString()
            Write-Verbose "Waiting operation complete ($($TimeDiff.TotalSeconds) seconds)."
        }

        if ($TimeDiff.TotalMinutes -ge $TimeoutInMinutes)
        {
            $error = "Operation timeout ($TimeoutInMinutes minutes)."

            $response = New-Object -TypeName PSObject `
                | Add-Member -PassThru -MemberType NoteProperty -Name StatusCode -Value 408 `
                | Add-Member -PassThru -MemberType NoteProperty -Name StatusDescription -Value "Request Timeout" `
                | Add-Member -PassThru -MemberType NoteProperty -Name Headers -Value $response.Headers `
                | Add-Member -PassThru -MemberType NoteProperty -Name Error -Value $error;
        }
    }

    return $response
}

#internal, helper method

function IsConnectorBlockingPolicyDefinitionValid
{
    param
    (
        [Parameter(Mandatory = $true)]
        [Object]$ConnectorBlockingDefinition
    )
    process
    {
        $alwaysEnabledConnectors = @(
            "shared_sharepointonline",
            "shared_sql",
            "shared_aci",
            "shared_approvals",
            "shared_arm",
            "shared_azuread",
            "shared_azureautomation",
            "shared_azureblob",
            "shared_azuredatafactory",
            "shared_azureeventgrid",
            "shared_azureeventgridpublish",
            "shared_azurefile",
            "shared_azureloganalyticsdatacollector",
            "shared_azurequeues",
            "shared_azuretables",
            "shared_cloudappsecurity",
            "shared_cognitiveservicescomputervision",
            "shared_cognitiveservicescontentmoderator",
            "shared_cognitiveservicescustomvision",
            "shared_cognitiveservicesqnamaker",
            "shared_cognitiveservicestextanalytics",
            "shared_commondataservice",
            "shared_commondataserviceforapps",
            "shared_documentdb",
            "shared_eventhubs",
            "shared_excel",
            "shared_faceapi",
            "shared_filesystem",
            "shared_flowmanagement",
            "shared_keyvault",
            "shared_luis",
            "shared_m365messagecenter",
            "shared_microsoft365compliance",
            "shared_microsoftflowforadmins",
            "shared_office365",
            "shared_office365groups",
            "shared_office365groupsmail",
            "shared_office365users",
            "shared_onenote",
            "shared_planner",
            "shared_powerappsforadmins",
            "shared_powerappsforappmakers",
            "shared_powerappsnotification",
            "shared_powerappsnotificationv2",
            "shared_powerbi",
            "shared_powerplatformforadmins",
            "shared_projectonline",
            "shared_servicebus",
            "shared_sftpwithssh",
            "shared_shifts",
            "shared_sqldw",
            "shared_teams",
            "shared_uiflow",
            "shared_wdatp",
            "shared_webcontents",
            "shared_wordonlinebusiness",
            "Http",
            "HttpRequestReceived",
            "HttpWebhook",
            "shared_microsoftforms",
            "shared_onedriveforbusiness"
        );

        foreach ($connectorSetting in $ConnectorBlockingDefinition.ConnectorSettings){
            $connectorIdsArray = $connectorSetting.id.Split("/");
            $connectorName = $connectorIdsArray[$connectorIdsArray.length-1];
            if ($alwaysEnabledConnectors -contains $connectorName){
                Write-Error "The connector $($connectorName) is always enabled and therefore not configurable as part of the Connector blocking policy. Please remove it from the policy and try again."
                return $false;
            }
        }

        return $true;
    }
}
# SIG # Begin signature block
# MIIoGwYJKoZIhvcNAQcCoIIoDDCCKAgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBh7/NtSDBnksbL
# E+v6ZCexpvnnCdGS537kTki+Z3FsfaCCDXYwggX0MIID3KADAgECAhMzAAAEhV6Z
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGfswghn3AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAASFXpnsDlkvzdcAAAAABIUwDQYJYIZIAWUDBAIB
# BQCggaAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIBkSRc2YMGv1VojwDT1p0CRW
# 9FtEoYRM+nfAmeu5hLzqMDQGCisGAQQBgjcCAQwxJjAkoBKAEABUAGUAcwB0AFMA
# aQBnAG6hDoAMaHR0cDovL3Rlc3QgMA0GCSqGSIb3DQEBAQUABIIBACEUS5IX8zQn
# qF/1QGP+ae5ZER/Bt5rk/sUAP0Z7CoYMGrtgyIc9wwU0hhh/crwMVe38aunUV5SI
# F4cn0CRZkl6VLSd95dAkmtKSHmbuvb1sz+GZ3DVq7wuoLmYQMpR9V8NkBQ7BqQgX
# hrd8ho6MQyzNbEFj/RVS99lC3EeHg0FacbUttZrIVMN9avKs/H+sVJQCX3OMhWDc
# fWyh63dTHuoknVGzMSznlkuKYCgOSsFgPOBgKt3oAq4iVvnbnB2M9iey+0ROF043
# CnEUVMmxGH3XIwGPAUqxv9cxrYeW7oGXEmjEVSpM9oAiLEaxSUNbnnwu0bPNB2Sb
# cNHLYd4bKQOhgheTMIIXjwYKKwYBBAGCNwMDATGCF38wghd7BgkqhkiG9w0BBwKg
# ghdsMIIXaAIBAzEPMA0GCWCGSAFlAwQCAQUAMIIBUgYLKoZIhvcNAQkQAQSgggFB
# BIIBPTCCATkCAQEGCisGAQQBhFkKAwEwMTANBglghkgBZQMEAgEFAAQgqgXT2mVB
# qpUYhwAXpd+YFNAzZm5wc/GtDbny8tz1SusCBmnBRzKEpxgTMjAyNjAzMjMxODM4
# NTAuNTA4WjAEgAIB9KCB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9u
# czEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkEwMDAtMDVFMC1EOTQ3MSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIR6TCCByAwggUIoAMC
# AQICEzMAAAIruwBQ/007mqEAAQAAAiswDQYJKoZIhvcNAQELBQAwfDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjYwMjE5MTk0MDExWhcNMjcwNTE3MTk0
# MDExWjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMG
# A1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNo
# aWVsZCBUU1MgRVNOOkEwMDAtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQg
# VGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAl95oujg97MlKkJuEKoJKyj23LCv0Md32HLS/PlTNbjmN26KIuRscGrk4EH+i
# RRyE06MUu4I6ipSvDhS8y+lE5dI8RCubeg7jnICV3b7rYpqE5TktAt5MiE1wQF6I
# /4KeoUUfc+lkYqdSrZIpW93SVwo0Kk/T9grro6/lc/K/mfow5dPY4v4nP+Bt+K95
# lcI7P/xp8fT7t9VfK1xYnDYgM8abm2sKW3fKan85Vk9r5xt5BfZejIkRG7yd1xy1
# MB0LIdLf060hcf7P8gqqSVmCeqApRu9Lb7BR9GkT/MAeHD/whWtiC75NuotznCQZ
# fqaiox00gcvZr8EzxA5Z83KNDbfEeqUj012YAbLHB4aCnwtFkJjs2NpHl2wJkU3G
# TMl8+b/wCW5qCNMtOwWs77eTZF3XRvUxK0FsLbBciCqxJQ4Fnx3gqE7tcLtnIg93
# Su9s93GtoM6BA8U9o/QVyFCmok803UD0bADGjt3VNM2hsDDJcLUicg4deGBIGaFL
# ub0vDLoDKnazY6Yci+ucioY6QFm4WJCBzv9LmY7vebT/M2TalyEYeLXX1hyTwE5/
# a/nMZMrodsdFS3X8dZZivV9zYx9DbYALOSQf8DpZMrrncZhU31lckay9+4rKTmfG
# jwBYL8kenDU5BqZBaN+SUY3IjZmYlOKk/VLcvleYLnRZNY8CAwEAAaOCAUkwggFF
# MB0GA1UdDgQWBBQ+Fo7kE1CW7W3d45r2ZLtBWdnlNjAfBgNVHSMEGDAWgBSfpxVd
# AF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIw
# UENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBo
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYG
# A1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0B
# AQsFAAOCAgEAzvwirHIhDPJK9X6h+E5X0+uhDaE48V8PNdKchKtD3a4C8H4E98ft
# YM+wkB7VHXr6jEOah8gy4ZuqU/ddQmJBjfuoPjFO3zGE6+nd0sYnicASKFpH0eIO
# 0orRszClOOuShGHo33XaFIKLwv8XEaWgCzuad/wNuPAcoSYjLbQUDQ7bE/x2ghcE
# RQlEW8v3/HNZJMvBfMZAlxc/vzLWeXdZVhY8DiNoHmR1qvV4oQzoHnuZ0tpKKOVe
# p/FxtttFE3r1X/qYJqSB+9Vyg1SGExhmSbOsj5Xydml6sNTBODUeqJDbGNz9TN9R
# +gzGEXyRjQTXqefeZFxod2MwN3AosoPo5iefIf307454CKblBXzg6Q4xcdInNWKC
# wDcYQhd0YUvamDOyuNDRISrIWLmgJCBtlwSmIoN6/9P29LI74wcLOeQGKJzJtwPK
# nF/+pPVX3NJr/XbaJx7lhnwNm/qhNqqQp4cxm3Qx6u4jkmRMNNZzbqQDH9XONZPS
# KE0Ns94sOsOGWaCzsoOEyjG6dZK6U+La4qf8t9Ar+ZIcqggzaml0KQZDmDjfC4La
# EN2plTl+4seY3a58f71MU1EooF761nS+1JPJKZktM7aNk6Mu2k+aAcwk734/Yifw
# TfxNb4RQZISQr2ez1b7DEp005pMdhWpdpVZM7bgCOOHw/7siyXWjEEswggdxMIIF
# WaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNy
# b3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAx
# ODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAy
# MDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL
# 1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5K
# Wv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTeg
# Cjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv62
# 6GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SH
# JMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss25
# 4o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/Nme
# Rd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afo
# mXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLi
# Mxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb
# 0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W2
# 9R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQF
# AgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1Ud
# DgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdM
# g30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJ
# KwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8w
# TTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
# dHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBK
# BggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9N
# aWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1V
# ffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1
# OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce57
# 32pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihV
# J9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZ
# UnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW
# 9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k
# +SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pF
# EUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L
# +DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1
# ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6
# CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIDTDCCAjQCAQEwgfmhgdGk
# gc4wgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNV
# BAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGll
# bGQgVFNTIEVTTjpBMDAwLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUACaw/dMpB6aP9ABm+5ZsL
# 7ArakTmggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkq
# hkiG9w0BAQsFAAIFAO1rxbAwIhgPMjAyNjAzMjMxMzU5MTJaGA8yMDI2MDMyNDEz
# NTkxMlowczA5BgorBgEEAYRZCgQBMSswKTAKAgUA7WvFsAIBADAGAgEAAgEDMAcC
# AQACAhCkMAoCBQDtbRcwAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkK
# AwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBADz9
# XdDaaYyJlkoaQZxz3LFKiseMdnjIyDzUukeMJ/Y2Hbxycg7DkA5ZBLkx7kEAnZ1u
# 31kYqBOwVDVVN0odSW5rmXC/mCEP39X8ROcrnaS3WsD7wCTmOmmKRVk2y3s8ttK9
# 6f3Zwux7k1+QWegynWSMMV+AjhwfsYR2cmqjP1000cuI8oz97Gsr5WtWB+iUvXNG
# Rpszf4b3hCY4FDdap1u3H91bdYyDQmJlb2xvrwQekA1o6Oef7K2uaJKjUn2oqFi4
# ZTRSw0kkJQD9/+qjx/vj7pH6qLf3JoWW70hu2QNc5nLteiTvWd6iMJ9/xCUrsS1k
# JmWujezlM2RguGJDEeoxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
# cCBQQ0EgMjAxMAITMwAAAiu7AFD/TTuaoQABAAACKzANBglghkgBZQMEAgEFAKCC
# AUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCD1
# Cx4uZO6CHogEYEafKyWd6PtcInOJDZ76ekaBhN1ejjCB+gYLKoZIhvcNAQkQAi8x
# geowgecwgeQwgb0EIHIOI/Q/kFftYA+M2OY+1Bx3ajBD6/WDAtPT2vFkv25SMIGY
# MIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIruwBQ/007
# mqEAAQAAAiswIgQgKitwtT06y0bt8B2T2xlSFUyOS1KAcWUrwgb7I4HXN+owDQYJ
# KoZIhvcNAQELBQAEggIAJu9q/GIO0F8ogucwWj3jIFPtHT8YfCDEf565+CjUBBCG
# VbFmP4Cd4LhKg7k9jT2HJlK0NJF57gnMbVDD0N7A2wNF/1nNyYLsmPbmOYh74kjX
# /9t9CUuSMQf67c0Z8K4on2PWBJ6TOZ6Y8w29TN4Wf6VAh4buiu8CiKndMYWtGagR
# D8gNt9z0YN7Q3Zy0VA9luGCpxgh9t8oS75yAcZBbc0v7KFl6QPDLdkE0leH3Zkof
# bHs2pT7sV0boBgJZCsReC6DF6BVOis3reVxQk+iTGhWggGWpJx7aR7W8xNOoEIBK
# TtQPA8NpsRnE6mxvc6XHq9YBrjuy+jrvkUjjOUgi2/I8MKAXq2HeHCdOx2a/8m4T
# dWk/rG4h6U3OzfkETAR+2XHfrkr+PkpnTQCcQzM8TRepiU1NkCmu/1pwwv/GBnK3
# Re9vxsCrdIEFFIuv0DIv/dh8uaYie58ZJsJzenhEmLl/kGEGlf5CHYfU3KBpDjlf
# wsLT2r5WhbC/mEhR+th8uA9v4AqszzE5SKBFkAzW0bBBKYTvlOP/6xufyMnWY+cd
# SkWnuvDpeP27en1/vHK84xAhjbUfAFwKRBA0qqeipuu9Zy7txD6Ypag2+HwnjG4g
# xuiCAvRdPW/0JIrlKg03LmITE8NmrsPRAxOmbc6sdFTRQG5Nej6hvNRsP7IGTYE=
# SIG # End signature block
