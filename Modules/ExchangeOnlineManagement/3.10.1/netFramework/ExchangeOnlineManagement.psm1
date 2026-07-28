# Import the REST module so that the EXO* cmdlets are present before Connect-ExchangeOnline in the powershell instance.
$RestModule = "Microsoft.Exchange.Management.RestApiClient.dll"
$RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestModule)
Import-Module $RestModulePath 

$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll"
$ExoPowershellModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule)
Import-Module $ExoPowershellModulePath

#Keep track of Execution status of last cmdlet
$global:EXO_LastExecutionStatus = $true;

############# Helper Functions Begin #############

    <#
    Get the ExchangeOnlineManagement module version.
    Same function is present in the autogen module. Both the codes should be kept in sync.
    #>
    function Get-ModuleVersion
    {
        try
        {
            # Return the already computed version info if available.
            if ($script:ModuleVersion -ne $null -and $script:ModuleVersion -ne '')
            {
                Write-Verbose "Returning precomputed version info: $script:ModuleVersion"
                return $script:ModuleVersion;
            }

            $exoModule = Get-Module ExchangeOnlineManagement
            
            # Check for ExchangeOnlineManagementBeta in case the psm1 is loaded directly
            if ($exoModule -eq $null)
            {
               $exoModule = (Get-Command -Name Connect-ExchangeOnline).Module
            }

            # Get the module version from the loaded module info.
            $script:ModuleVersion = $exoModule.Version.ToString()

            # Look for prerelease information from the corresponding module manifest.
            $exoModuleRoot = (Get-Item $exoModule.Path).Directory.Parent.FullName

            $exoModuleManifestPath = Join-Path -Path $exoModuleRoot -ChildPath ExchangeOnlineManagement.psd1
            $isExoModuleManifestPathValid = Test-Path -Path $exoModuleManifestPath
            if ($isExoModuleManifestPathValid -ne $true)
            {
                # Could be a local debug build import for testing. Skip extracting prerelease info for those.
                Write-Verbose "Module manifest path invalid, path: $exoModuleManifestPath, skipping extracting prerelease info"
                return $script:ModuleVersion
            }

            $exoModuleManifestContent = Get-Content -Path $exoModuleManifestPath
            $preReleaseInfo = $exoModuleManifestContent -match "Prerelease = '(.*)'"
            if ($preReleaseInfo -ne $null)
            {
                $script:ModuleVersion = "{0}-{1}" -f $exoModule.Version.ToString(),$preReleaseInfo[0].Split('=')[1].Trim().Trim("'")
            }

            Write-Verbose "Computed version info: $script:ModuleVersion"
            return $script:ModuleVersion
        }
        catch
        {
            return [string]::Empty
        }
    }

    <#
    .Synopsis Validates a given Uri
    #>
    function Test-Uri
    {
        [CmdletBinding()]
        [OutputType([bool])]
        Param
        (
            # Uri to be validated
            [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
            [string]
            $UriString
        )

        [Uri]$uri = $UriString -as [Uri]

        $uri.AbsoluteUri -ne $null -and $uri.Scheme -eq 'https'
    }

    <#
    .Synopsis Is Cloud Shell Environment
    #>
    function global:IsCloudShellEnvironment()
    {
        return [Microsoft.Exchange.Management.AdminApiProvider.Utility]::IsCloudShellEnvironment();
    }

    <#
    .SYNOPSIS Extract organization name from UserPrincipalName
    #>
    function Get-OrgNameFromUPN
    {
        param([string] $UPN)
        $fields = $UPN -split '@'
        return $fields[-1]
    }

    <#
    .SYNOPSIS Get the command from the given module
    #>
    function global:Get-WrappedCommand
    {
        param(
        [string] $CommandName,
        [string] $ModuleName,
        [string] $CommandType)

        $cmd = (Get-Module $moduleName).ExportedFunctions[$CommandName]
        return $cmd
    }

    <#
    .Synopsis Writes a message to the console in Yellow.
    Call this method with caution since it uses Write-Host internally which cannot be suppressed by the user.
    #>
    function Write-Message
    {
        param([string] $message)
        Write-Host $message -ForegroundColor Yellow
    }

############# Helper Functions End #############

###### Begin Main ######

$EOPConnectionInProgress = $false
function Connect-ExchangeOnline 
{
    [CmdletBinding()]
    param(

        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = '',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = '',

        # Exchange Environment name
        [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment] $ExchangeEnvironmentName = 'O365Default',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # Prefix 
        [string] $Prefix = '',

        # Show Banner of Exchange cmdlets Mapping and recent updates
        [switch] $ShowBanner = $true,

        #Cmdlets to Import for rps cmdlets , by default it would bring all
        [string[]] $CommandName = @("*"),

        #The way the output objects would be printed on the console
        [string[]] $FormatTypeName = @("*"),

        # Use to Skip Exchange Format file loading into current shell.
        [switch] $SkipLoadingFormatData = $false,

        # Use to skip downloading and loading the help files into the current connection.
        [switch] $SkipLoadingCmdletHelp = $true,

        # Use to enable downloading and loading the help files into the current connection.
        [switch] $LoadCmdletHelp = $false,

        # Externally provided access token
        [string] $AccessToken = '',

        # Client certificate to sign the temp module
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $SigningCertificate = $null,

        # switch to disable WAM
        [switch] $DisableWAM = $false,

        # Temp Base Path
        [string] $EXOModuleBasePath  = [System.IO.Path]::GetTempPath()
    )
    DynamicParam
    {
        if (($isCloudShell = IsCloudShellEnvironment) -eq $false)
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # User Principal Name or email address of the user
            $UserPrincipalName = New-Object System.Management.Automation.RuntimeDefinedParameter('UserPrincipalName', [string], $attributeCollection)
            $UserPrincipalName.Value = ''

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            # Certificate
            $Certificate = New-Object System.Management.Automation.RuntimeDefinedParameter('Certificate', [System.Security.Cryptography.X509Certificates.X509Certificate2], $attributeCollection)
            $Certificate.Value = $null

            # Certificate Path 
            $CertificateFilePath = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateFilePath', [string], $attributeCollection)
            $CertificateFilePath.Value = ''

            # Certificate Password
            $CertificatePassword = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificatePassword', [System.Security.SecureString], $attributeCollection)
            $CertificatePassword.Value = $null

            # Certificate Thumbprint
            $CertificateThumbprint = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateThumbprint', [string], $attributeCollection)
            $CertificateThumbprint.Value = ''

            # Application Id
            $AppId = New-Object System.Management.Automation.RuntimeDefinedParameter('AppId', [string], $attributeCollection)
            $AppId.Value = ''

            # Organization
            $Organization = New-Object System.Management.Automation.RuntimeDefinedParameter('Organization', [string], $attributeCollection)
            $Organization.Value = ''

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")

            # Create a new attribute and valiate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

            # Switch to use Managed Identity flow. 
            $ManagedIdentity = New-Object System.Management.Automation.RuntimeDefinedParameter('ManagedIdentity', [switch], $attributeCollection)
            $ManagedIdentity.Value = $false

            # ManagedIdentityAccountId to be used in case of User Assigned Managed Identity flow
            $ManagedIdentityAccountId = New-Object System.Management.Automation.RuntimeDefinedParameter('ManagedIdentityAccountId', [string], $attributeCollection)
            $ManagedIdentityAccountId.Value = ''

# EXO params start

            # Switch to track perfomance 
            $TrackPerformance = New-Object System.Management.Automation.RuntimeDefinedParameter('TrackPerformance', [bool], $attributeCollection)
            $TrackPerformance.Value = $false

            # Flag to enable or disable showing the number of objects written
            $ShowProgress = New-Object System.Management.Automation.RuntimeDefinedParameter('ShowProgress', [bool], $attributeCollection)
            $ShowProgress.Value = $false

            # Switch to enable/disable Multi-threading in the EXO cmdlets
            $UseMultithreading = New-Object System.Management.Automation.RuntimeDefinedParameter('UseMultithreading', [bool], $attributeCollection)
            $UseMultithreading.Value = $true

            # Pagesize Param
            $PageSize = New-Object System.Management.Automation.RuntimeDefinedParameter('PageSize', [uint32], $attributeCollection)
            $PageSize.Value = 1000

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            # Switch to CmdInline parameters
            $InlineCredential = New-Object System.Management.Automation.RuntimeDefinedParameter('InlineCredential', [switch], $attributeCollection)
            $InlineCredential.Value = $false

# EXO params end
            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('Certificate', $Certificate)
            $paramDictionary.Add('CertificateFilePath', $CertificateFilePath)
            $paramDictionary.Add('CertificatePassword', $CertificatePassword)
            $paramDictionary.Add('AppId', $AppId)
            $paramDictionary.Add('Organization', $Organization)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            $paramDictionary.Add('TrackPerformance', $TrackPerformance)
            $paramDictionary.Add('ShowProgress', $ShowProgress)
            $paramDictionary.Add('UseMultithreading', $UseMultithreading)
            $paramDictionary.Add('PageSize', $PageSize)
            $paramDictionary.Add('ManagedIdentity', $ManagedIdentity)
            $paramDictionary.Add('ManagedIdentityAccountId', $ManagedIdentityAccountId)
            if($PSEdition -eq 'Core')
            {
                $paramDictionary.Add('Device', $Device)
                $paramDictionary.Add('InlineCredential', $InlineCredential);
                # We do not want to expose certificate thumprint in Linux as it is not feasible there.
                if($IsWindows)
                {
                    $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
                }
            }
            else 
            {
                $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
            }

            return $paramDictionary
        }
        else
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")

            # Create a new attribute and validate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

            # Switch to CmdInline parameters
            $InlineCredential = New-Object System.Management.Automation.RuntimeDefinedParameter('InlineCredential', [switch], $attributeCollection)
            $InlineCredential.Value = $false

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Device', $Device)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('InlineCredential', $InlineCredential)
            return $paramDictionary
        }
    }
    process {
        $global:EXO_LastExecutionStatus = $true;
        $EnableSearchOnlySession = $false
        if (Test-Path Variable:Script:EnableSearchOnlySession)
        {
            $var = Get-Variable -Name "EnableSearchOnlySession" -Scope Script -ErrorAction SilentlyContinue
            if ($null -ne $var) {
                $EnableSearchOnlySession = $var.Value
            }
        }
        $startTime = Get-Date

        # Validate parameters
        if (($ConnectionUri -ne '') -and (-not (Test-Uri $ConnectionUri)))
        {
            $global:EXO_LastExecutionStatus = $false;
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        if (($AzureADAuthorizationEndpointUri -ne '') -and (-not (Test-Uri $AzureADAuthorizationEndpointUri)))
        {
            $global:EXO_LastExecutionStatus = $false;
            throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"
        }
        if (($Prefix -ne ''))
        {
            if ($Prefix -notmatch '^[a-z0-9]+$') 
            {
                $global:EXO_LastExecutionStatus = $false;
                throw "Use of any special characters in the Prefix string is not supported."
            }
            if ($Prefix -eq 'EXO') 
            {
                $global:EXO_LastExecutionStatus = $false;
                throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
            }
        }

        # Keep track of error count at beginning.
        $errorCountAtStart = $global:Error.Count;
        try
        {
            $moduleVersion = Get-ModuleVersion
            Write-Verbose "ModuleVersion: $moduleVersion"

            # Generate a ConnectionId to use in all logs and to send in all server calls.
            $connectionContextID = [System.Guid]::NewGuid()

            $cmdletLogger = New-CmdletLogger -ExoModuleVersion $moduleVersion -LogDirectoryPath $LogDirectoryPath.Value -EnableErrorReporting:$EnableErrorReporting.Value -ConnectionId $connectionContextID
            $logFilePath = $cmdletLogger.GetCurrentLogFilePath()
            
            if ($EnableErrorReporting.Value -eq $true)
            {
                Write-Message ("Writing cmdlet logs to " + $logFilePath)
            }

            $cmdletLogger.InitLog($connectionContextID)
            $cmdletLogger.LogStartTime($connectionContextID, $startTime)
            if ($EOPConnectionInProgress -eq $true)
            {
                $cmdletLogger.LogCmdletName($connectionContextID, "Connect-IPPSSession");
            }
            else
            {
                $cmdletLogger.LogCmdletName($connectionContextID, "Connect-ExchangeOnline");
            }
            $cmdletLogger.LogCmdletParameters($connectionContextID, $PSBoundParameters);

            if ($isCloudShell -eq $false)
            {
                $ConnectionContext = Get-ConnectionContext -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri `
                -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value `
                -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring `
                -DelegatedOrg $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value `
                -CertificatePassword $CertificatePassword.Value -CertificateThumbprint $CertificateThumbprint.Value -AppId $AppId.Value `
                -Organization $Organization.Value -Device:$Device.Value -InlineCredential:$InlineCredential.Value -CommandName $CommandName `
                -FormatTypeName $FormatTypeName -Prefix $Prefix -PageSize $PageSize.Value -ExoModuleVersion:$moduleVersion -Logger $cmdletLogger `
                -ConnectionId $connectionContextID -EnableErrorReporting:$EnableErrorReporting.Value `
                -ManagedIdentity:$ManagedIdentity.Value -ManagedIdentityAccountId $ManagedIdentityAccountId.Value -AccessToken $AccessToken -DisableWAM:$DisableWAM -LogDirectoryPath $LogDirectoryPath.Value -EnableSearchOnlySession $EnableSearchOnlySession -EXOModuleBasePath $EXOModuleBasePath
            }
            else
            {
                $ConnectionContext = Get-ConnectionContext -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri `
                -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -Credential $Credential.Value -PSSessionOption $PSSessionOption `
                -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -InlineCredential:$InlineCredential.Value `
                -DelegatedOrg $DelegatedOrganization -CommandName $CommandName -FormatTypeName $FormatTypeName -Prefix $Prefix -ExoModuleVersion:$moduleVersion `
                -Logger $cmdletLogger -ConnectionId $connectionContextID -EnableErrorReporting:$EnableErrorReporting.Value -AccessToken $AccessToken -DisableWAM:$DisableWAM -LogDirectoryPath $LogDirectoryPath.Value -EnableSearchOnlySession $EnableSearchOnlySession -EXOModuleBasePath $EXOModuleBasePath

            }

            if ($isCloudShell -eq $false)
            {
                $global:_EXO_EnableErrorReporting = $EnableErrorReporting.Value;
            }

            if ($ShowBanner -eq $true)
            {
                try
                {
                    $BannerContent = Get-EXOBanner -ConnectionContext:$ConnectionContext
                    Write-Host -ForegroundColor Yellow $BannerContent
                }
                catch
                {
                    Write-Verbose "Failed to fetch banner content from server. Reason: $_"
                    $cmdletLogger.LogGenericError($connectionContextID, $_);
                }
            }

            if (($ConnectionUri -ne '') -and ($AzureADAuthorizationEndpointUri -eq ''))
            {
                Write-Information "Using ConnectionUri:'$ConnectionUri', in the environment:'$ExchangeEnvironmentName'."
            }
            if (($AzureADAuthorizationEndpointUri -ne '') -and ($ConnectionUri -eq ''))
            {
                Write-Information "Using AzureADAuthorizationEndpointUri:'$AzureADAuthorizationEndpointUri', in the environment:'$ExchangeEnvironmentName'."
            }

            $ImportedModuleName = '';
            $LogModuleDirectoryPath = [System.IO.Path]::GetTempPath();

            

            # Download the new web based EXOModule
            if ($SigningCertificate -ne $null)
            {
                $ImportedModule = New-EXOModule -ConnectionContext $ConnectionContext -SkipLoadingFormatData:$SkipLoadingFormatData -SigningCertificate $SigningCertificate;
            }
            else
            {
                $ImportedModule = New-EXOModule -ConnectionContext $ConnectionContext -SkipLoadingFormatData:$SkipLoadingFormatData;
            }
            if ($null -ne $ImportedModule)
            {
                $ImportedModuleName = $ImportedModule.Name;
                $LogModuleDirectoryPath = $ImportedModule.ModuleBase

                Write-Verbose "AutoGen EXOModule created at  $($ImportedModule.ModuleBase)"

                if ($LoadCmdletHelp -eq $true -and $HelpFileNames -ne $null -and $HelpFileNames -is [array] -and $HelpFileNames.Count -gt 0)
                {
                    Get-HelpFiles -HelpFileNames $HelpFileNames -ConnectionContext $ConnectionContext -ImportedModule $ImportedModule -EnableErrorReporting:$EnableErrorReporting.Value
                }
                else
                {
                    $cmdletLogger.LogGenericInfo($connectionContextID, "Skipping cmdlet help data");
                }
            }
            else
            {
                $global:EXO_LastExecutionStatus = $false;
                throw "Module could not be correctly formed. Please run Connect-ExchangeOnline again."
            }
            

            # If we are configured to collect telemetry, add telemetry wrappers in case of an RPS connection
            if ($EnableErrorReporting.Value -eq $true)
            {

                $endTime = Get-Date
                $cmdletLogger.LogEndTime($connectionContextID, $endTime);
                $cmdletLogger.CommitLog($connectionContextID);

                if ($EOPConnectionInProgress -eq $false)
                {
                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }
            }
            else 
            {
                if ($EOPConnectionInProgress -eq $false)
                {
                    # Set the AppSettings disabling the logging
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -EnableErrorReporting $false
                }
            }
        }
        catch
        {
            # If telemetry is enabled, log errors generated from this cmdlet also.
            # If telemetry is not enabled, calls to cmdletLogger will be a no-op.
            $errorCountAtProcessEnd = $global:Error.Count 
            $numErrorRecordsToConsider = $errorCountAtProcessEnd - $errorCountAtStart
            for ($i = 0 ; $i -lt $numErrorRecordsToConsider ; $i++)
            {
                $cmdletLogger.LogGenericError($connectionContextID, $global:Error[$i]);
            }

            $cmdletLogger.CommitLog($connectionContextID);

            if ($EnableErrorReporting.Value -eq $true)
            {
                if ($global:_EXO_TelemetryFilePath -eq $null)
                {
                    $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath -LogDirectoryPath $LogDirectoryPath.Value

                    # Import the REST module
                    $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                    $RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestPowershellModule);
                    Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings;

                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ConnectionUri $ConnectionUri -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }

                # Log errors which are encountered during Connect-ExchangeOnline execution. 
                Write-Message ("Writing Connect-ExchangeOnline error log to " + $global:_EXO_TelemetryFilePath)
                Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Connect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
            }

            $global:EXO_LastExecutionStatus = $false;

            if ($_.Exception -ne $null)
            {
                # Connect-ExchangeOnline Failed, Remove ConnectionContext from Map.
                if ([Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::RemoveConnectionContextUsingConnectionId($connectionContextID))
                {
                    Write-Verbose "ConnectionContext Removed"
                }

                if ($_.Exception.InnerException -ne $null)
                {
                    throw $_.Exception.InnerException;
                }
                else
                {
                    throw $_.Exception;
                }
            }
            else
            {
                throw $_;
            }
        }
    }
}

function Connect-IPPSSession
{
    [CmdletBinding()]
    param(
        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations',

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,

        # Prefix 
        [string] $Prefix = '',

        #Cmdlets to Import, by default it would bring all
        [string[]] $CommandName = @("*"),

        #The way the output objects would be printed on the console
        [string[]] $FormatTypeName = @("*"),

        # Show Banner of scc cmdlets Mapping and recent updates
        [switch] $ShowBanner = $true,

        # switch to disable WAM
        [switch] $DisableWAM = $false,

        # Externally provided access token
        [string] $AccessToken = '',
        
        # Switch to using OBO Token
        [switch] $EnableSearchOnlySession = $false,

        # TempBasePath
        [string] $EXOModuleBasePath = [System.IO.Path]::GetTempPath()
    )
    DynamicParam
    {
        if (($isCloudShell = IsCloudShellEnvironment) -eq $false)
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # User Principal Name or email address of the user
            $UserPrincipalName = New-Object System.Management.Automation.RuntimeDefinedParameter('UserPrincipalName', [string], $attributeCollection)
            $UserPrincipalName.Value = ''

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            # Certificate
            $Certificate = New-Object System.Management.Automation.RuntimeDefinedParameter('Certificate', [System.Security.Cryptography.X509Certificates.X509Certificate2], $attributeCollection)
            $Certificate.Value = $null

            # Certificate Path 
            $CertificateFilePath = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateFilePath', [string], $attributeCollection)
            $CertificateFilePath.Value = ''

            # Certificate Password
            $CertificatePassword = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificatePassword', [System.Security.SecureString], $attributeCollection)
            $CertificatePassword.Value = $null

            # Certificate Thumbprint
            $CertificateThumbprint = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateThumbprint', [string], $attributeCollection)
            $CertificateThumbprint.Value = ''

            # Application Id
            $AppId = New-Object System.Management.Automation.RuntimeDefinedParameter('AppId', [string], $attributeCollection)
            $AppId.Value = ''

            # Organization
            $Organization = New-Object System.Management.Automation.RuntimeDefinedParameter('Organization', [string], $attributeCollection)
            $Organization.Value = ''

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")
            # Create a new attribute and validate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('Certificate', $Certificate)
            $paramDictionary.Add('CertificateFilePath', $CertificateFilePath)
            $paramDictionary.Add('CertificatePassword', $CertificatePassword)
            $paramDictionary.Add('AppId', $AppId)
            $paramDictionary.Add('Organization', $Organization)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            if($PSEdition -eq 'Core')
            {
                # We do not want to expose certificate thumprint in Linux as it is not feasible there.
                if($IsWindows)
                {
                    $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
                }
            }
            else 
            {
                $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
            }

            return $paramDictionary
        }
        else
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")
            # Create a new attribute and validate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)
            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Device', $Device)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            return $paramDictionary
        }
    }
    process 
    {
        try
        {
            $EOPConnectionInProgress = $true
            Set-Variable -Name "EnableSearchOnlySession" -Value $EnableSearchOnlySession.IsPresent -Scope Script
            if ($isCloudShell -eq $false)
            {
                $LogLevelValue = [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All
                if($LogLevel.Value)
                {
                    $LogLevelValue = $LogLevel.Value
                }
                $certThumbprint = $CertificateThumbprint.Value
                # Will pass CertificateThumbprint if it is not null or not empty
                if($certThumbprint)
                {
                    Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value -CertificatePassword $CertificatePassword.Value -CertificateThumbprint $certThumbprint -AppId $AppId.Value -Organization $Organization.Value -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -DisableWAM:$DisableWAM -EnableErrorReporting:$EnableErrorReporting.Value -LogLevel $LogLevelValue -LogDirectoryPath $LogDirectoryPath.Value -EXOModuleBasePath $EXOModuleBasePath
                }
                else
                {
                    Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value -CertificatePassword $CertificatePassword.Value -AppId $AppId.Value -Organization $Organization.Value -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -DisableWAM:$DisableWAM -AccessToken $AccessToken -EnableErrorReporting:$EnableErrorReporting.Value -LogLevel $LogLevelValue -LogDirectoryPath $LogDirectoryPath.Value -EXOModuleBasePath $EXOModuleBasePath
                }
            }
            else
            {
                Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $PSSessionOption -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -DisableWAM:$DisableWAM -AccessToken $AccessToken -EnableErrorReporting:$EnableErrorReporting.Value -LogLevel $LogLevelValue -LogDirectoryPath $LogDirectoryPath.Value -EXOModuleBasePath $EXOModuleBasePath
            }
        }
        finally
        {
            $EOPConnectionInProgress = $false
        }
    }
}

function Disconnect-ExchangeOnline
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High', DefaultParameterSetName='DefaultParameterSet')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='ConnectionId', ValueFromPipelineByPropertyName=$true)]
        [string[]] $ConnectionId,
        [Parameter(Mandatory=$true, ParameterSetName='ModulePrefix')]
        [string[]] $ModulePrefix
    )

    process
    {
        $global:EXO_LastExecutionStatus = $true;
        $disconnectConfirmationMessage = ""
        Switch ($PSCmdlet.ParameterSetName)
        {
            'ConnectionId'
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageByConnectionId($ConnectionId)
                break
            }
            'ModulePrefix'
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageByModulePrefix($ModulePrefix)
                break
            }
            Default
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageWithInbuilt()
            }
        }
	
        if ($PSCmdlet.ShouldProcess(
            $disconnectConfirmationMessage,
            "Press(Y/y/A/a) if you want to continue.",
            $disconnectConfirmationMessage))
        {

            # Keep track of error count at beginning.
            $errorCountAtStart = $global:Error.Count;
            $startTime = Get-Date

            try
            {
                # Get all the connection contexts so that the logger can be initialized.
                $connectionContexts = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetAllConnectionContexts()
                $disconnectCmdletId = [System.Guid]::NewGuid().ToString()

                # denotes if any of the connections is an RPS session.
                # This is used to Push-EXOTelemetryRecord in case any RPS connections are present.
                $rpsConnectionWithErrorReportingExists = $false
                
                foreach ($context in $connectionContexts)
                {
                    $context.Logger.InitLog($disconnectCmdletId);
                    $context.Logger.LogStartTime($disconnectCmdletId, $startTime);
                    $context.Logger.LogCmdletName($disconnectCmdletId, "Disconnect-ExchangeOnline");
                    if ($context.IsRpsSession -and $context.ErrorReportingEnabled)
                    {
                        $rpsConnectionWithErrorReportingExists = $true
                    }
                }

                # Import the module once more to ensure that Test-ActiveToken is present
                $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll";
                $ModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule);
                Import-Module $ModulePath -Cmdlet Clear-ActiveToken;

                $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}

                if ($existingPSSession.count -gt 0) 
                {
                    for ($index = 0; $index -lt $existingPSSession.count; $index++)
                    {
                        $session = $existingPSSession[$index]
                        Remove-PSSession -session $session

                        Write-Information "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"

                        # Remove any active access token from the cache
                        # If the connectionId of the session being cleared is equal to AppSettings.ConnectionId, this means connection to EXO cmdlets will break.
                        if ($session.ConnectionContext.ConnectionId -ieq [Microsoft.Exchange.Management.AdminApiProvider.AppSettings]::ConnectionId)
                        {
                            Clear-ActiveToken -TokenProvider $session.TokenProvider -IsSessionUsedByInbuiltCmdlets:$true
                        }
                        else
                        {
                            Clear-ActiveToken -TokenProvider $session.TokenProvider -IsSessionUsedByInbuiltCmdlets:$false
                        }

                        # Remove any previous modules loaded because of the current PSSession
                        if ($session.PreviousModuleName -ne $null)
                        {
                            if ((Get-Module $session.PreviousModuleName).Count -ne 0)
                            {
                                $null = Remove-Module -Name $session.PreviousModuleName -ErrorAction SilentlyContinue
                            }

                            $session.PreviousModuleName = $null
                        }

                        # Remove any leaked module in case of removal of broken session object
                        if ($session.CurrentModuleName -ne $null)
                        {
                            if ((Get-Module $session.CurrentModuleName).Count -ne 0)
                            {
                                $null = Remove-Module -Name $session.CurrentModuleName -ErrorAction SilentlyContinue
                            }
                        }
                    }
                }

                $modulesToRemove = $null
                Switch ($PSCmdlet.ParameterSetName)
                {
                    'ConnectionId'
                    {
                        # Call GetModulesToRemoveByConnectionId in this scenario
                        $modulesToRemove = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetModulesToRemoveByConnectionId($ConnectionId)
                        break
                    }
                    'ModulePrefix'
                    {
                        # Call GetModulesToRemoveByModulePrefix in this scenario
                        $modulesToRemove = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetModulesToRemoveByModulePrefix($ModulePrefix)
                        break
                    }
                    Default
                    {
                        # Remove all the AutoREST modules from this instance of powershell if created
                        $existingAutoRESTModules = Get-Module "tmpEXO_*"
                        foreach ($module in $existingAutoRESTModules)
                        {  
                            $null = Remove-Module -Name $module -ErrorAction SilentlyContinue
                        }

                        # The below call to remove all connection contexts could be removed as we already have an OnRemove event hooked to the module. Work Item 3461604 to investigate this
                        # Remove all ConnectionContexts
                        # this internally clears all the active tokens in ConnectionContexts
                        [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::RemoveAllConnectionContexts()
                    }
                }

                if ($modulesToRemove -ne $null -and $modulesToRemove.Count -gt 0)
                {
                    $null = Remove-Module $modulesToRemove -ErrorAction SilentlyContinue
                }

                Write-Information "Disconnected successfully !"

                # Remove all the Wrapped modules from this instance of powershell if created
                $existingWrappedModules = Get-Module "EXOCmdletWrapper-*"
                foreach ($module in $existingWrappedModules)
                {
                    $null = Remove-Module -Name $module -ErrorAction SilentlyContinue
                }

                if ($rpsConnectionWithErrorReportingExists)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid
                }
            }
            catch
            {
                # If telemetry is enabled for any of the connections, log errors generated from this cmdlet also. 
                $errorCountAtProcessEnd = $global:Error.Count
                $global:EXO_LastExecutionStatus = $false;
                
                $endTime = Get-Date
                foreach ($context in $connectionContexts)
                {
                    $numErrorRecordsToConsider = $errorCountAtProcessEnd - $errorCountAtStart
                    for ($i = 0 ; $i -lt $numErrorRecordsToConsider ; $i++)
                    {
                        $context.Logger.LogGenericError($disconnectCmdletId, $global:Error[$i]);
                    }

                    $context.Logger.LogEndTime($disconnectCmdletId, $endTime);
                    $context.Logger.CommitLog($disconnectCmdletId);
                    $logFilePath = $context.Logger.GetCurrentLogFilePath();
                }

                if ($rpsConnectionWithErrorReportingExists)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    # Log errors which are encountered during Disconnect-ExchangeOnline execution. 
                    Write-Message ("Writing Disconnect-ExchangeOnline errors to " + $global:_EXO_TelemetryFilePath)

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
                }

                throw $_
            }

            $endTime = Get-Date
            foreach ($context in $connectionContexts)
            {
                if ($context.Logger -ne $null)
                {
                    $context.Logger.LogEndTime($disconnectCmdletId, $endTime);
                    $context.Logger.CommitLog($disconnectCmdletId);
                }
            }
        }
    }
}

# SIG # Begin signature block
# MIInbgYJKoZIhvcNAQcCoIInXzCCJ1sCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCANB1d3lhVWz4gv
# 5WWuzFPaC9iOQovbzFaO7WgzbOZnLaCCDMkwggYEMIID7KADAgECAhMzAAACHPrN
# xZvoL37EAAAAAAIcMA0GCSqGSIb3DQEBCwUAMFcxCzAJBgNVBAYTAlVTMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBD
# b2RlIFNpZ25pbmcgUENBIDIwMjQwHhcNMjYwNDE2MTg1OTQxWhcNMjcwNDE1MTg1
# OTQxWjB0MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYD
# VQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQDVsZfgOKmM31HPfoWOoNEiw0SlCiIxUMC0I9NMWbucKOw/e9lP
# oAoehQVu6SG65V4EPzrYsnBnFPNoi4/HoOdjhz1qkrEt4I6tEcxXU6oOeY9zGveC
# /3iBeuhLYxM3M/PkcUoebF+Nednm8OkdSPoDu8imViHPQq/8CQUu0WRR4rE+dMRf
# rpVqfmNi2qWCX94T4MsepijGVkwE//tJg0ryAiYdHT34LSnlG/RSBZmQRGWZ5g8j
# qnKjRParSqMft1gvjuUTVgtWNZfgcLFSK5Wa0myrq8OPcgTGGsRgun+tnSS+IxDT
# xVsAPH1OzvPjwomguByhUe/OcvUN0D5Wmp7xAgMBAAGjggGqMIIBpjAOBgNVHQ8B
# Af8EBAMCB4AwHwYDVR0lBBgwFgYKKwYBBAGCN0wIAQYIKwYBBQUHAwMwHQYDVR0O
# BBYEFNoH7a2YDjOSwpkp6DHcmUS7J+0yMFQGA1UdEQRNMEukSTBHMS0wKwYDVQQL
# EyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxFjAUBgNVBAUT
# DTIzMDAxMis1MDc1NjkwHwYDVR0jBBgwFoAUf1k/VCHarU/vBeXmo9ctBpQSCDEw
# YAYDVR0fBFkwVzBVoFOgUYZPaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
# cy9jcmwvTWljcm9zb2Z0JTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDI0LmNy
# bDBtBggrBgEFBQcBAQRhMF8wXQYIKwYBBQUHMAKGUWh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwQ29kZSUyMFNpZ25pbmcl
# MjBQQ0ElMjAyMDI0LmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4IC
# AQAUnEqhaRXe0T3hIJjvdQErEkrA/7bByjn6t5IArODkkRjzkYwtKMc2yYj2quaN
# rLutWw2YZcngKPy1b71YyDJQTy4NDRwaSh9Tw5thrk3NmcPrAHia5vtcBJ1CgtKK
# 7mQbIcQ22d/N3813ayCDDFewu1+jsZmX+r/aTEqaOM4TVxVtRSkuCy8nAXKuChOK
# Li/zA4XuH8iEYqIsj2YoNaeSxVmeGiERXpKdo3dDmYi0kO5w2D8VS4c3+9h6gElY
# BaAAg/dYErBg27qT3vv0zRDJhJufvCNylA8S7/+8H5E/PV5cng6na9VV/w9OV3qu
# uND6zdGa2EX38Glp50F9AIQk3p2xXmcvorDeM4XJ7UlWYBi6g80J1SSOQnInCYFE
# msfUNn3+1AaTJKSJL83quKArTac2pKhu0Yzzzrzo6HrsRiQKzpnRBb1/dMa6P3hz
# 75XbMRBctNsFhZC07WCmjExdLg2eHW5uV0TY8D5+6wozJf7vF3+WHkYPO85Z+BC6
# U4FkNbYNycZ9cE4j1tXRdyDCfml6c0HWPHjNVDObrv9lKt3qUqFpX38VCqVCyNOO
# 1UcXfQiVjJw32U2WUKZjt/neJKHEBsm9kFsLuWzkQ53+qcaSaytmsCnk2gOglrlD
# 5d3kKyvvAw+rzm0lT8K38P6PLxfZQHhu4W8dV7Av8N2ZmDCCBr0wggSloAMCAQIC
# EzMAAAA5O7Y3Gb8GHWcAAAAAADkwDQYJKoZIhvcNAQEMBQAwgYgxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDExMB4XDTI0MDgwODIwNTQxOFoX
# DTM2MDMyMjIyMTMwNFowVzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQ
# Q0EgMjAyNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANgBnB7jOMeq
# lRYHNa265v4IY9fH8TKhemHfPINe1gpLaV3dhg324WwH06LcHbpnsBukCDNitryo
# 0dtS/EW6I/yEL/bLSY8hKpbfQuWusBPr9qazYcDxCW/qnjb5JsI1s8bNOg3bVATv
# QVL4tcf03aTycsz8QeCdM0l/yHRObJ9QqazM1r6VPEOJ7LL+uEEb73w6QCuhs89a
# 1uv1zerOYMnsneRRwCbpyW11IcggU0cRKDDq1pjVJzIbIF6+oiXXbReOsgeI8zu1
# FyQfK0fVkaya8SmVHQ/tOf23mZ4W9k0Ri22QW9p3UgSC5OUDktKxxcCmGL6tXLfO
# GSWHIIV4YrTJTT6PNty5REojHJuZHArkF9VnHTERWoTjAzfI3kP+5b4alUdhgAZ7
# ttOu1bVnXfHaqPYl2rPs20ji03LOVWsh/radgE17es5hL+t6lV0eVHrVhsssROWJ
# uz2MXMCt7iw7lFPG9LXKGjsmonn2gotGdHIuEg5JnJMJVmixd5LRlkmgYRZKzhxS
# CwyoGIq0PhaA7Y+VPct5pCHkijcIIDm0nlkK+0KyepolcqGm0T/GYQRMhHJlGOOm
# VQop36wUVUYklUy++vDWeEgEo4s7hxN6mIbf2MSIQ/iIfMZgJxC69oukMUXCrOC3
# SkE/xIkgpfl22MM1itkZ35nNXkMolU1lAgMBAAGjggFOMIIBSjAOBgNVHQ8BAf8E
# BAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFH9ZP1Qh2q1P7wXl5qPX
# LQaUEggxMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMA8GA1UdEwEB/wQFMAMB
# Af8wHwYDVR0jBBgwFoAUci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0fBFMwUTBP
# oE2gS4ZJaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMv
# TWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcBAQRSMFAw
# TgYIKwYBBQUHMAKGQmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMv
# TWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNydDANBgkqhkiG9w0BAQwFAAOC
# AgEAFJQfOChP7onn6fLIMKrSlN1WYKwDFgAddymOUO3FrM8d7B/W/iQ6DxXsDn7D
# 5W4wMwYeLystcEqfkjz4NURRgazyMu5yRzQh4LqjA4tStTcJh1opExo7nn5PuPBY
# nbu0+THSuVHTe0VTTPVhily/piFrDo3axQ9P4C+Ol5yet+2gTfekICS5xS+cYfSI
# vgn0JksVBVMYVI5QFu/qhnLhsEFEUzG8fvv0hjgkO+lkpV9ty6GkN4vdnd7ya6Q6
# aR9y34aiM1qmxaxBi6OUnyNl6fkuun/diTFnYDLTppOkr/mg5WSfCiDVMNCxtj4w
# PKC5OmHm1DQIt/MNokbbH3UGsFP1QbzsLocuSqLCvH09Io3fDPTmscR9Y75G4qX7
# RTX8AdBPo0I6OEojf39zuFZt0qOHm65YWQE69cZM2ueE1MB05dNNgHK9gTE7zKvK
# /fg8B2qjW88MT/WF5V5uvZGtqa9FSL2RazArA+rDPuf6JGYz4HpgMZHB4S6szWSK
# YBv0VisCzfxgeU+dquXW9bd0auYlOB58DPcOYKdc3Se94g+xL4pcEhbB54JOgAkw
# YTu/9dLeH2pDqeJZAABVDWRQCaXfO5LgyKwKCLYXpigrZYCjUSBcr+Ve8PFWMhVT
# Ql0v4q8J/AUmQN5W4n101cY2L4A7GTQG1h32HHAvfQESWP0xghn7MIIZ9wIBATBu
# MFcxCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# KDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMjQCEzMAAAIc
# +s3Fm+gvfsQAAAAAAhwwDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIELnVufEpNJxhIRgGZbDKwkKYEYFOdZvuQlJdRT4gKXcMEIGCisG
# AQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEBBQAEggEADbY0AgENPjybUc5cZptk
# 8xUMvcb2Jx9YOd0HdEs87kowjRUtshmbAz38TFdEANsTsK9AJ+Z8ed3OLXO/NarY
# r15QFdoohHCqhveM+oRPAhDHliKulMaODT0pWzP2L8D6iXxcg/DFqnqB6ppjg6Kw
# D0zGIgPuVlf/aYejys1DYrlhm3mHolcIWcLQHoR7pS05QvdLrnGMZo8a+e7CK2DP
# 3bT2WkyXqJL6UjwT/BTNI3JX8XYAhWtMDOAmnFKd+vsHnutjYPoVnw31LRign+wu
# QBbKxZs+50gP/aCalvumkTvxPHhazsVZNghbmIdLUTeAf6vcWi8fxYgcZC7e6irQ
# naGCF60wghepBgorBgEEAYI3AwMBMYIXmTCCF5UGCSqGSIb3DQEHAqCCF4YwgheC
# AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFaBgsqhkiG9w0BCRABBKCCAUkEggFFMIIB
# QQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCDNZSdPU8q4eH5dqzm5
# kcUfLd9nU+YisHsKiwKN6s/MNgIGajlP/wCgGBMyMDI2MDcwMjAyMDYwNC4xMjJa
# MASAAgH0oIHZpIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0
# ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjozNjA1LTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCEfswggcoMIIFEKAD
# AgECAhMzAAACE7BDNWbPr5XoAAEAAAITMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDgxNDE4NDgxN1oXDTI2MTExMzE4
# NDgxN1owgdMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTAr
# BgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUG
# A1UECxMeblNoaWVsZCBUU1MgRVNOOjM2MDUtMDVFMC1EOTQ3MSUwIwYDVQQDExxN
# aWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEA9Jl64LoZxDINSFgz+9KS5Ozv5m548ePVzc9RXWe4T4/Mplfg
# a4eq12RGdp5cVvnjde5vxfq2ax/jnu7vUW4rZN4mOUm5vh+kcYsQlYQ53FwgIB3n
# EjcQHomrG3mZe/ozjFSAr6JbglKtIeAySPzAcFzyAer5lLNUHBEvQMM8BOjMyapC
# vh0xsg4xKFcVEJQLKEfCGBffMZI/amutHFb3CUTZ7aVpG2KHEFUNlZ1vwMKvxXTP
# RDnbwPGzyyqJJznfsLNHQ4vXt2ttS1PeCoGI0hN1Peq8yGsIXM9oocwC06DGNSM/
# 4LAx2uKvwmUn6NwLc0+tmvny6w28rZLejskRfnVWofEv1mWY0jHUnHrwSGBS8gVP
# 9gcBs6P5g0OpJPMfxdUkHXRkcMPPW0hIP8NbW8W5Sup8HuwnSKbjpyAlGBUdM/V5
# rZb0sZmkn714r6ULGK+cLLAN6R3FhX6N0nj64F27LTK2BbS0pJZaXjo0eDNz1Qcx
# eIFLUgF+RBsLYDn8E8cCkexK8Nlt3Gi9zJf55w6UfTZ+kwTMxMqFxh7+Tfx7+aBO
# bZ+nx961AtiqAy7zVV69o/LWRdKPZdvZn9ESyGbTnPfjkBERv22prSlETlRwzP6b
# mEVOKWLWVwxuwh7bUWUuUb1cj93zvttQYGQat5E9ALLJNmlvLKCskB7raLsCAwEA
# AaOCAUkwggFFMB0GA1UdDgQWBBQTnhBKx+FryphQWMRipH49sMFAOjAfBgNVHSME
# GDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1l
# LVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsG
# AQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01p
# Y3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMB
# Af8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDAN
# BgkqhkiG9w0BAQsFAAOCAgEAgmxaJrGqQ2D6UJhZ6Ql2SZFOaNuGbW3LzB+ES+l2
# BB1MJtBRSFdi/hVY33NpxsJQhQ5TLVp0DXYOkIoPQc17rH+IVhemO8jCt+U6I1TI
# w6cR7c+tEo/Jjp6EqEU1c4/mraMjgHhQ+raC/OUAm98A1r4bIPHtsBmLROGmeE5X
# LIFaBIZWHvh2COXITKObXVd5wGtJ1dZZdwaHACXF506jta+uoUdyzAeuNlTPLTrZ
# 8nyhxGwk9Vh6eiDQ7CQMWSSa8DJS9PUXjeoi9vTdS7ZMXqu+tv6Qz3xtoBF5+YFK
# 4uE+miGs90Fxm0VK2lWrmFhjkRl5zyoHOdwG7spNYkDomCPNWIudUQmQYKpt/Hss
# pfcb+xpnWIDQdMzgE8pj1vpwLgWEnH7LtT4dZCeoDo9PK40RxBD8kKJ769ngkEwf
# wCD2EX/MQk79eIvOhpnH12GuVByvaKZk5XZvqtPONNwr8q/qA3877IuWwWgnaeX+
# prpw0dZ/QLtbGGVrgP+TRQjt+2dcZA5P3X4LwANhiPsy0Ol4XCdj7OxBLFvOzsCP
# DPaVnkp+dfDFG+NOBir7aqTJ68622pymg1V+6gc/1RvxC/wgvYyG033ecJqv0On0
# ZRNYr+i/OkwgA3HP1aLD0aHrEpw6lt0263iRkCvrcdcOW8w3jC8TJuaGWyC2S9jE
# jzgwggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqGSIb3DQEB
# CwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYD
# VQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAe
# Fw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA5OGm
# TOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4XYDP/XE/H
# ZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTzxXb1hlDc
# wUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7uhp7M62A
# W36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlwaQ3KNi1w
# jjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedGbsoy1cCG
# MFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXNxF803RKJ
# 1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03dJQcNIIP
# 8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9ahhaYQFz
# ymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5UPkLiWHz
# NgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReTwDwV2xo3
# xwgVGD94q0W29R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZMBIGCSsG
# AQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8RjUpzxD/
# LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAEVTBTMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAwwCgYIKwYB
# BQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8G
# A1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQw
# VgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9j
# cmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUF
# BwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZIhvcNAQEL
# BQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9MTO1OdfC
# cTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2LpypglYAA7AF
# vonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbzaN9l
# 9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8DJ6LGYnn
# 8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA0j3mSj5m
# O0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1RtnWN0SCyx
# TkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjcZxkoJLo4
# S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFmPWn9
# y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJC4822rpM
# +Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328y+l7vzhw
# RNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIDVjCCAj4C
# AQEwggEBoYHZpIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0
# ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjozNjA1LTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIa
# AxUAmBE8SCjxgjacmy8/VEdk7NxpR6aggYMwgYCkfjB8MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1T
# dGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQsFAAIFAO3vq04wIhgPMjAyNjA3MDEx
# NTA1NTBaGA8yMDI2MDcwMjE1MDU1MFowdDA6BgorBgEEAYRZCgQBMSwwKjAKAgUA
# 7e+rTgIBADAHAgEAAgImNjAHAgEAAgIT1zAKAgUA7fD8zgIBADA2BgorBgEEAYRZ
# CgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0G
# CSqGSIb3DQEBCwUAA4IBAQAyBThr637w6jN/eBXAlhmkVkuUtd+sgca4T1abRNiM
# d9ZakWBnpm261lotZEqaWscBDU8i1ZGhezSSCzKXtyksiDQ4nBN5c/zXpqi8N9fK
# +N/squOrBZikbzfvyLfBk/qOp6a/NVpxb796YE5ACqJy+M+jvfvtlMUVGG+eCoyi
# ExCzzsKHFMBCQDneZmLgLZmYi+rwC37sTpqH+xzedEGs6WEwKnd7s/l6c9Z7hFYa
# Z+7apORVQH5MItPNNUdTN8IHG3wWNBqSocMboxClzSg4iIvEnKMFbM3/mGwYm+i7
# OYOHuWRToMSUCQX5UOqyX/tFfEdASmjiYap5AO2R0D1oMYIEDTCCBAkCAQEwgZMw
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAITsEM1Zs+vlegAAQAA
# AhMwDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRAB
# BDAvBgkqhkiG9w0BCQQxIgQgl1kERVt7bOQXIMJbT21IdFB2wKSLvVUYmuAF5oIz
# ArQwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCDM4QltFIUz8J4DjAzP4nVo
# dZvQxYGleUIfp86Oa5xYaDCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAACE7BDNWbPr5XoAAEAAAITMCIEICe6vLrSAFoLvEVfgLmX8oo+
# kfVL7R6xqQ3Ii0up/xpkMA0GCSqGSIb3DQEBCwUABIICAOfDqrHAmJpMIgLWl7Ih
# EewWaRzTRbNaxOfKtBA1JIOVvtbF5vkEUDxZSLKu91lyLS86mJWztHMnVILPcWUY
# NCcTH8Rdu+8M4vCWxV6MS6b7eBiidT9wVEDtpCRwYYurjBh8Vw+fzmsoJqzfDP6h
# eEOG4roFGC6/PytIhYu9sVXqBNBV4D79i1U+OXtX7jJIFPr3v9rvYqIrbBexNm92
# sMpi6LaF8YpYQmOHsfjlV1fFD453e6YAg+5F0CRVu+999x7ttbtQqMqL/iPQLJb9
# cMYGXEYlnXhPtJ4UlKEju7CLsIVfPnIDuSzKRPeqbtqACT3Cgex2Wu8HM42YWX99
# dTbCY761yBq+x0OLCxGlDonYIOlrtQ/y1343VH4WPbdxQjJH/Vvnk2TD18QtfzEi
# jkCSkfgjnvFOkVn1BPu3160t2nKppkSHXGz4CbUuEGXvt1faVahY2k1Sn65N5oGC
# 7rUgjBYY2J09YnSDwl8oH1VHumUAyDvZdNaereDqLg890nWyr46B9JycAsmjqU9s
# TUvBwAQBNikbhxsL1KDRKqdQQJCRN7p278prIt05iKLJPa7wVn2TkmMD/mmWWTUM
# 8jsogV0V4lK6fa+wRIjuxVRgiE0Bz7d5UCxmr93x2Mx0aybiPdoBSYzwvoEWpC9k
# T511F99ZV7GDliwV6iXZV76I
# SIG # End signature block
