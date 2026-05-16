@{

# Script module or binary module file associated with this manifest.
RootModule = 'Microsoft.PowerApps.Administration.Powershell.psm1'

# Version number of this module.
ModuleVersion = '2.0.217'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = '1c40b0da-ee6a-4226-9a3d-e60092e1daae'

# Author of this module
Author = 'Microsoft Common Data Service Team'

# Company or vendor of this module
CompanyName = 'Microsoft'

# Copyright statement for this module
Copyright = '© 2020 Microsoft Corporation. All rights reserved'

# Description of the functionality provided by this module
Description = 'PowerShell interface for Microsoft PowerApps and Flow Administrative features'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '3.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
PowerShellHostVersion = '1.0'

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
DotNetFrameworkVersion = '4.0.0.0'

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
#RequiredModules = @(@{ModuleName = "Microsoft.PowerApps.RestClientModule"; ModuleVersion = "1.0"; Guid = "04800678-e13e-4b41-8d46-424e707ea733"})
#RequiredModules = @(@{ModuleName = "Microsoft.PowerApps.RestClientModule"; ModuleVersion = "1.0"; Guid = "04800678-e13e-4b41-8d46-424e707ea733"})

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
#NestedModules = @('Microsoft.PowerApps.AuthModule', 'Microsoft.PowerApps.RestClientModule')

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @(
    'New-AdminPowerAppCdsDatabase', `
    'Get-AdminPowerAppCdsDatabaseLanguages', `
    'Get-AdminPowerAppCdsDatabaseCurrencies', `
    'Get-AdminPowerAppEnvironmentLocations', `
    'Get-AdminPowerAppCdsDatabaseTemplates', `
    'New-AdminPowerAppEnvironment', `
    'Set-AdminPowerAppEnvironmentDisplayName', `
    'Set-AdminPowerAppEnvironmentBackupRetentionPeriod', `
    'Set-AdminPowerAppEnvironmentRuntimeState', `
	'Set-AdminPowerAppEnvironmentGovernanceConfiguration', `
    'Get-AdminPowerAppEnvironment', `
    'Get-AdminPowerAppSoftDeletedEnvironment', `
    'Get-AdminPowerAppOperationStatus', `
    'Remove-AdminPowerAppEnvironment', `
    'Recover-AdminPowerAppEnvironment', `
    'Reset-PowerAppEnvironment', `
    'Get-AdminPowerAppEnvironmentRoleAssignment', `
    'Set-AdminPowerAppEnvironmentRoleAssignment', `
    'Remove-AdminPowerAppEnvironmentRoleAssignment', `
    'Get-AdminPowerApp', `
    'Remove-AdminPowerApp', `
    'Get-AdminPowerAppRoleAssignment', `
    'Remove-AdminPowerAppRoleAssignment', `
    'Set-AdminPowerAppRoleAssignment', `
    'Set-AdminPowerAppOwner', `
    'Get-AdminFlow', `
    'Add-PowerAppsCustomBrandingAssets', `
    'Enable-AdminFlow', `
    'Disable-AdminFlow', `
    'Remove-AdminFlow', `
    'Restore-AdminFlow', `
    'Remove-AdminFlowApprovals', `
    'Set-AdminFlowOwnerRole', `
    'Remove-AdminFlowOwnerRole', `
    'Get-AdminFlowOwnerRole', `
    'Get-AdminPowerAppConnector', `
    'Get-AdminFlowEnvironmentCmkStatus', `
    'Get-AdminFlowEncryptedByMicrosoftKey', `
    'Get-AdminFlowWithHttpAction', `
    'Get-AdminFlowWithMigratingTriggerUrl', `
    'Get-DesktopFlowModules', `
    'Add-AdminFlowsToSolution', `
    'Start-EUDBMigration', `
    'Cancel-EUDBMigration', `
    'Get-EUDBMigrationStatus', `
    'Add-AdminFlowPowerAppContext', `
    'Remove-AdminFlowPowerAppContext', `
    'Get-AdminFlowAtRiskOfSuspension', `
    'Get-AdminPowerAppConnectorAction', `
    'Get-AdminPowerAppConnectorRoleAssignment', `
    'Set-AdminPowerAppConnectorRoleAssignment', `
    'Remove-AdminPowerAppConnectorRoleAssignment', `
    'Get-AdminGenerateDataverseEnforceReport', `
    'Remove-AdminPowerAppConnector', `
    'Get-AdminPowerAppConnection', `
    'Remove-AdminPowerAppConnection', `
    'Get-AdminPowerAppConnectionRoleAssignment', `
    'Set-AdminPowerAppConnectionRoleAssignment', `
    'Remove-AdminPowerAppConnectionRoleAssignment', `
    'Get-AdminPowerAppsUserDetails', `
    'Get-AdminFlowUserDetails', `
    'Remove-AdminFlowUserDetails', `
    'Set-AdminPowerAppAsFeatured', `
    'Clear-AdminPowerAppAsFeatured', `
    'Set-AdminPowerAppAsHero', `
    'Clear-AdminPowerAppAsHero', `
    'Set-AppAsUnquarantined', `
    'Set-AppAsQuarantined', `
	'Get-AppQuarantineState', `
    'Set-AdminPowerAppApisToBypassConsent', `
    'Clear-AdminPowerAppApisToBypassConsent', `
    'Get-AdminPowerAppConditionalAccessAuthenticationContextIds', `
    'Set-AdminPowerAppConditionalAccessAuthenticationContextIds', `
    'Remove-AdminPowerAppConditionalAccessAuthenticationContextIds', `
    'Get-AdminDlpPolicy', `
    'New-AdminDlpPolicy', `
    'Remove-AdminDlpPolicy', `
    'Set-AdminDlpPolicy', `
    'Add-ConnectorToBusinessDataGroup', `
    'Remove-ConnectorFromBusinessDataGroup', `
    'Get-AdminPowerAppConnectionReferences', `
    'Add-CustomConnectorToPolicy', `
    'Add-ConnectorsToPolicy', `
    'Remove-CustomConnectorFromPolicy', `
    'Remove-LegacyCDSDatabase', `
    'Get-AdminDeletedPowerAppsList', `
    'Get-AdminRecoverDeletedPowerApp', `
    'Add-AdminAllowedThirdPartyApps', `
    'Get-AdminAllowedThirdPartyApps', `
    'Remove-AdminAllowedThirdPartyApps', `
    'Set-PowerPlatformGenerativeAiSettings', `
    'Set-AdminPowerAppEnvironmentMakerAnalyticsSettings', `
    #from Rest and Auth Module Helpers
    'Select-CurrentEnvironment', `
    'Add-PowerAppsAccount', `
    'Remove-PowerAppsAccount',`
    'Test-PowerAppsAccount', `
    'Get-TenantDetailsFromGraph', `
    'Get-UsersOrGroupsFromGraph', `
    'Get-JwtToken', `
    'ReplaceMacro', `
    'Set-TenantSettings', `
    'Get-TenantSettings', `
    'Get-AdminPowerAppTenantConsumedQuota', `
    'InvokeApi', `
    'InvokeApiNoParseContent', `
    'Add-AdminPowerAppsSyncUser', `
    'Remove-AllowedConsentPlans', `
    'Add-AllowedConsentPlans', `
    'Get-AllowedConsentPlans', `
    'Get-AdminPowerAppCdsAdditionalNotificationEmails', `
    'Set-AdminPowerAppCdsAdditionalNotificationEmails', `
    'Get-AdminPowerAppLicenses', `
    'Set-AdminPowerAppDesiredLogicalName' `
    # DLP policy Version 1 APIs
    'Get-DlpPolicy', `
    'New-DlpPolicy', `
    'Remove-DlpPolicy', `
    'Set-DlpPolicy', `
    # URL patterns Version 1 APIs
    'Get-PowerAppTenantUrlPatterns', `
    'New-PowerAppTenantUrlPatterns', `
    'Remove-PowerAppTenantUrlPatterns', `
    'Get-PowerAppPolicyUrlPatterns', `
    'New-PowerAppPolicyUrlPatterns', `
    'Remove-PowerAppPolicyUrlPatterns', `
    # Dlp policy connector configurations Version 1 APIs
    'Get-PowerAppDlpPolicyConnectorConfigurations', `
    'New-PowerAppDlpPolicyConnectorConfigurations', `
    'Remove-PowerAppDlpPolicyConnectorConfigurations', `
    'Set-PowerAppDlpPolicyConnectorConfigurations', `
    # Copy/Backup/Restore APIs
    'Copy-PowerAppEnvironment', `
    'Backup-PowerAppEnvironment', `
    'Get-PowerAppEnvironmentBackups', `
    'Restore-PowerAppEnvironment', `
    'Remove-PowerAppEnvironmentBackup', `
    # Tenant To Tenant Migration APIs
    'TenantToTenant-PrepareMigration', `
    'TenantToTenant-GetMigrationStatus', `
    'TenantToTenant-MigratePowerAppEnvironment', `
    'TenantToTenant-SubmitMigrationRequest',`
    'TenantToTenant-ManageMigrationRequest',`
    'TenantToTenant-ViewMigrationRequest',`
    'TenantToTenant-ViewApprovalRequest',`
    'TenantToTenant-DeleteMigrationRequest',`
    'TenantToTenant-UploadUserMappingFile',`
    # Generate Resource Storage API
    'GenerateResourceStorage-PowerAppEnvironment', `
    # ManagementApp APIs
    'Get-PowerAppManagementApp', `
    'Get-PowerAppManagementApps', `
    'New-PowerAppManagementApp', `
    'Remove-PowerAppManagementApp', `
    # Environment Keywords
    'Get-AdminPowerAppSharepointFormEnvironment', `
    'Set-AdminPowerAppSharepointFormEnvironment', `
    'Reset-AdminPowerAppSharepointFormEnvironment', `
    # Protection key APIs
    'Get-PowerAppGenerateProtectionKey', `
    'Get-PowerAppRetrieveTenantProtectionKey', `
    'Get-PowerAppRetrieveAvailableTenantProtectionKeys', `
    'New-PowerAppImportProtectionKey', `
    'Set-PowerAppProtectionStatus', `
    'Set-PowerAppTenantProtectionKey', `
    'Set-PowerAppLockAllEnvironments', `
    'Set-PowerAppUnlockEnvironment', `
    # Tenant isolation APIs
    'Get-PowerAppTenantIsolationPolicy', `
    'Set-PowerAppTenantIsolationPolicy', `
    'Get-PowerAppTenantIsolationOperationStatus', `
	# Dlp/Governance Error Settings APIs
	'Get-PowerAppDlpErrorSettings', `
	'New-PowerAppDlpErrorSettings', `
	'Set-PowerAppDlpErrorSettings', `
	'Remove-PowerAppDlpErrorSettings', `
	'Get-GovernanceErrorSettings', `
	'New-GovernanceErrorSettings', `
	'Set-GovernanceErrorSettings', `
	'Remove-GovernanceErrorSettings', `
    # Dlp policy exempt resources Version 1 APIs
	'Get-PowerAppDlpPolicyExemptResources', `
	'New-PowerAppDlpPolicyExemptResources', `
	'Remove-PowerAppDlpPolicyExemptResources', `
	'Set-PowerAppDlpPolicyExemptResources', `
	# virtual connector Route
	'Get-AdminVirtualConnectors', `
    # Dlp Enforcement on Connections APIs
    'Start-DLPEnforcementOnConnectionsInTenant', `
    'Start-DLPEnforcementOnConnectionsInEnvironment', `
    # Dlp Connector blocking APIs
    'Get-PowerAppDlpConnectorBlockingPolicies', `
    'Get-PowerAppDlpConnectorBlockingPolicy', `
    'New-PowerAppDlpConnectorBlockingPolicy', `
    'Set-PowerAppDlpConnectorBlockingPolicy', `
    'Remove-PowerAppDlpConnectorBlockingPolicy', `
    # Admin power platform requests consumption
    'Get-AdminPowerPlatformRequestsConsumptionOfFlows',
    'Get-AdminNonZoneRedundantFlows',
    'Get-AdminEnableZoneRedundancyPreflightCheck',
    'Start-AdminFlowEnableZoneRedundancy',
    'Get-AdminEnableZoneRedundancyStatus'
)

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
# CmdletsToExport = @()

# Variables to export from this module
# VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
# AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
ModuleList = @("Microsoft.PowerApps.Administration.PowerShell" )

# List of all files packaged with this module
# When included they are automatically loaded which can pull the files by name from uncontrolled locations.
FileList = @(
    "Microsoft.PowerApps.Administration.PowerShell.psm1", `
    "Microsoft.PowerApps.Administration.PowerShell.psd1", `
    "Microsoft.PowerApps.AuthModule.psm1", `
    "Microsoft.PowerApps.RestClientModule.psm1"
)

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
         LicenseUri = 'https://aka.ms/powerappspowershellterms'

        # A URL to the main website for this project.
         ProjectUri = 'https://docs.microsoft.com/en-us/powerapps/administrator/powerapps-powershell'

        # A URL to an icon representing this module.
         IconUri = 'https://connectoricons-prod.azureedge.net/powerplatformforadmins/icon_1.0.1056.1255.png'

        # ReleaseNotes of this module
        ReleaseNotes = '

Current Release:
2.0.217
    Update example for the Cancel-EUDBMigration cmdlet
    Added new apis for support getting flows with trigger urls that require migration.
    Get-AdminFlowWithMigratingTriggerUrl

2.0.209
    Adding function to upload user mapping file in single command.
    TenantToTenant-UploadUserMappingFile

2.0.208
    Update examples for the Power Automate EUDB migration

2.0.206
    Add support for the EUDB cancel experience.
    Cancel-EUDBMigration
	Added new apis for support getting flows with http action in an environment.
    Get-AdminFlowWithHttpAction

2.0.205
    Add support for migrating flows into EUDB compliant resource groups
	Start-EUDBMigration
	Get-EUDBMigrationStatus

2.0.203
    Added new apis for support getting flows that remain encrypted by Microsoft Managed Key in an environment.
    Get-AdminFlowEncryptedByMicrosoftKey

2.0.198
    Update required params in TenantToTenant-PrepareMigration and TenantToTenant-MigratePowerAppEnvironment. We need MigrationID to initiate a Tenant To Tenant Migrations.
    Added new APIs to support the submission and approval process for Tenant to Tenant Migrations.
    Removed TenantToTenant-GetStatus Command which is not supported anymore.
    TenantToTenant-ViewApprovalRequest
    TenantToTenant-GetMigrationStatus
    TenantToTenant-DeleteMigrationRequest

2.0.197
    Added new apis for Submitting and Approving Tenant To Tenant Migrations.
    TenantToTenant-SubmitMigrationRequest
    TenantToTenant-ManageMigrationRequest
    TenantToTenant-ViewMigrationRequest

2.0.170
    Add Support for Tenant To Tenant Migration and for generating resource storage.
    TenantToTenant-SubmitMigrationRequest
    TenantToTenant-ManageMigrationRequest
    TenantToTenant-ViewMigrationRequest

2.0.169
    Add Support for Tenant To Tenant Migration and for generating resource storage.
     TenantToTenant-GetStatus
     TenantToTenant-MigratePowerAppEnvironment
     TenantToTenant-PrepareMigration
     GenerateResourceStorage-PowerAppEnvironment

2.0.168
	Add support for Tenant To Tenant Migration and for generating resource storage.
     TenantToTenant-GetStatus
     TenantToTenant-MigratePowerAppEnvironment
     TenantToTenant-PrepareMigration
     GenerateResourceStorage-PowerAppEnvironment

2.0.167
    Add support for migrating non-solution flows to solution as admin.
        Add-AdminFlowsToSolution
    Support ''CreatedBy'' filter which was removed in 2.0.166. We can support this filter now as backend API has been fixed to return the CreatedBy metadata.

2.0.166
    Update Get-AdminFlow to use v2 route for fetching flows. This v2 route returns less information about flows.
    Response will contain the following flow properties: Id, display name, created time, last modified time, state and workflow entity Id.
    BREAKING CHANGE: Removed support for ''CreatedBy'' filter as v2 route does not return the creator of the flow.

2.0.165
    Add support for setting the background operations state on an environment.
        Set-AdminPowerAppEnvironmentRuntimeState

2.0.163
    Fix polling support in Set-AdminPowerAppEnvironmentRuntimeState
        Set-AdminPowerAppEnvironmentRuntimeState        

2.0.157
    Added new Apis to get environment cmk status in PowerAutomate. 
        Get-AdminFlowEnvironmentCmkStatus

2.0.156
    Add support for setting backup retention period on an environment.
        Set-AdminPowerAppEnvironmentBackupRetentionPeriod

2.0.155
    Add flow at risk of suspension function.
        Get-AdminFlowAtRiskOfSuspension

2.0.155
    Add support for Developer environment provisioning and TemplateMetadata parameter for Dataverse provisioning.

2.0.154
    Update license uri link

2.0.153
	Add new governance error settings functions
		Get-GovernanceErrorSettings
		New-GovernanceErrorSettings
		Set-GovernanceErrorSettings
		Remove-GovernanceErrorSettingss
	PowerAppDlpErrorSettings will be deprecated in a future version.

2.0.150
    Remove warning from Set-AdminPowerAppEnvironmentGovernanceConfiguration'
    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = 'PowerApp'

}

# SIG # Begin signature block
# MIIoHwYJKoZIhvcNAQcCoIIoEDCCKAwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCJ9liFvKPV/Sg0
# +uhACR16CRQXQkwkUjDzQmYjCC0ww6CCDXYwggX0MIID3KADAgECAhMzAAAEhV6Z
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGf8wghn7AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAASFXpnsDlkvzdcAAAAABIUwDQYJYIZIAWUDBAIB
# BQCggaAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIIf7bRs9tkZc62x3KTtPL4tM
# qfOnf59GqHpR7DndzEawMDQGCisGAQQBgjcCAQwxJjAkoBKAEABUAGUAcwB0AFMA
# aQBnAG6hDoAMaHR0cDovL3Rlc3QgMA0GCSqGSIb3DQEBAQUABIIBAK9/fHwtoLL4
# VQllugJTwoTSmaNH/KvJRkv6yDWkKhbh0/ybOqPdXIGYgUSCnjG7HeRfnchb/ATN
# B6iKj28dGRFfFCAXMFUbU9E7JZ2D0Fms9RTru2mkf0jJu1WjBOaEWCHeqSxYe2RA
# rtiJDHwVA/PJLlGaw46ekTo+r+SVUJU7IbWMZMQtdbK2E9qIM06ptL4bGc+hY5sz
# K162XYZ4vbbyjXK6WmAofgUo1qVpL6kqsRcwE9jz+VGyJmZLEAggELxgArjfui/p
# QbTsm+5pOB7pPSpHidZIeuu2+9VLLpOHn3uEEqVoiIjFltppOBElSbiUmPl0do6o
# O30ZyYu/3OOhgheXMIIXkwYKKwYBBAGCNwMDATGCF4Mwghd/BgkqhkiG9w0BBwKg
# ghdwMIIXbAIBAzEPMA0GCWCGSAFlAwQCAQUAMIIBUgYLKoZIhvcNAQkQAQSgggFB
# BIIBPTCCATkCAQEGCisGAQQBhFkKAwEwMTANBglghkgBZQMEAgEFAAQg8tiRFIfJ
# jSfgnFo6HL5gXANXG2EjLZQMi1/9EYtL8aACBmm4bST5phgTMjAyNjAzMjMxODM4
# NTAuMzYyWjAEgAIB9KCB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9u
# czEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjMzMDMtMDVFMC1EOTQ3MSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIR7TCCByAwggUIoAMC
# AQICEzMAAAIhM8A1+9IPIaQAAQAAAiEwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjYwMjE5MTkzOTU0WhcNMjcwNTE3MTkz
# OTU0WjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMG
# A1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNo
# aWVsZCBUU1MgRVNOOjMzMDMtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQg
# VGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA23EwAqlNWL0aHMli9jy/X8n//lC7Nqiu1NWmbEZw2Up5Qq+yu44AN3hQhCS+
# QWe3VEwtA3mXqX/mQvuxxGweCHc5iX0AFAxRXq6mOVUx5kLz9lwN5VkhY++NInXB
# lB4JT+R/z2wiVOxgB1j9h3XAo3cdZWAKNAPsyyO8cJ00HjMjl19tdhIOFJgzzyYM
# XUzMOlhVVrAT1kQYuYA4sctrPu0fAA5OZWwQRQweYdAo6zViDe7ggMxeYO7a6y/J
# 1yCqddJo/UcYXBkPrZYbelSL3coEVU1BncxQdv5wbyakPZMcRZbUEk+9HxHceE8m
# iqMP3+fgUoeM+P/X+zVyFVUy5//JHCQH0ahZka6xbdyCm8u1a85mLqEFg9JZjRbR
# kOewayZD6zxQD3pNQC7XG2+xR950Kb4vJ4M/zBV//nJ5jRVhVNvVVS5swfV7y2cW
# 2L5HnrbdJoeZX7XnjdqxMFMq3ayrn8/YdkuqW2rXvgtodNgq18EpGtMens6U5hpC
# CSxbdubm/1GFzS3R3bMRg+hH3JDiKCWLJuDEvRf70qizRyvPSNL0ywZ4EBKeiyBZ
# CDWp0U9z7Tcd6TSkSiUQC3Oi+poVuIS+Ghy++Paj30O9reagDJucYimDICdlmp4n
# USzbiNudSSDe62mngP9r29FxZGXCG00daX0BrHKOFNIObY8CAwEAAaOCAUkwggFF
# MB0GA1UdDgQWBBTmIyLOamuqX7qrj8sitRU6+UAwpzAfBgNVHSMEGDAWgBSfpxVd
# AF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIw
# UENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBo
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYG
# A1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0B
# AQsFAAOCAgEAOCP925HZ08Q9qxbptzBfMVSnRJKIQDm23j60PtH6+W0Ypo8/bFZC
# k/+4HI+DjHApUmBviHV+jKdxGLCx1n478H8xaHmRPsk23QY/9VR2UEbgpsOkKnlQ
# k28Np50u5wcZ1nfaGV2z1KahGsB+Q6l0GlhYEfQOCllSvyL11QzI9T5TwhEtT9ya
# JzW3YZJJM+PaybijpuW+3vwR/JaKgJlzl0XNtssVlUzFqxKeKbJZr/Hk+1aGPF/4
# 3SmEz1RF7H5i21RXKszLgfLxRn1MlrFkTkvMIKu5UGH1nGKoezcpqAE1/sFmCt81
# hu2kXIjxlAM8513X/mh7SFp0CzWuRxZkl5ImpN30rqa1mGYh4bmIxNeoa6AKXAR6
# ZvvEv5DaoZvVo0F/tgcZ2L/iXo8upak4vHywS0tOvVl1cP6bX+SFfhbWJd+Br1aH
# oN9VKFJlVWXtUg1CZJvXQ13PJf6gQ2IgCE9ggrD08rfVwPSVbh8XT+t5+wob1gDv
# +O0Ebgg7FJRSaFsMgcJe43mKWkVTLULdIriTBho4BGiV9UP9o/LF1Eb03Hixww/Y
# qVrdPdmQ1jEHIg0ZoRzRTl9XZ4wb5P5NVDHIPfe4+aGM5wJ0qSb5YP+AT92lRNIf
# 2B9ioLCm1ODV2RwIyV49kpaqNQtdeQhuqgWWhZDPFurz2Qpuap0nszowggdxMIIF
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
# CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIDUDCCAjgCAQEwgfmhgdGk
# gc4wgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNV
# BAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGll
# bGQgVFNTIEVTTjozMzAzLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUAC2xIGWZ8mB1ydQxm+Xxo
# 6ZV6bbmggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkq
# hkiG9w0BAQsFAAIFAO1rfK8wIhgPMjAyNjAzMjMwODQ3NDNaGA8yMDI2MDMyNDA4
# NDc0M1owdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA7Wt8rwIBADAKAgEAAgIbIwIB
# /zAHAgEAAgISjTAKAgUA7WzOLwIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEE
# AYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBCwUAA4IB
# AQCR3Ll9QKZkcDPk4N5YTM8XYVP7LbLkIl1VVI4Bg24L/OmY7pQXhH8ue4ksRj85
# n5gRC79VfrDmbmiwO65vjT4mPRIFdu8bTpHadBkgYBJXlxmBYBE70ZjoK0ATPjoz
# G+9CfM/b0NOOpmLjEX0L0/IugH3+Y467b+3hFtcG79vnytWY7eiTfxyKeAO8I2Dm
# FzdYOO4+CDis288EzwElPCNB2g5Rxpamx5kAPF0l+5MbWcm+ozRk52tYEnfmyvrN
# Y/pP4DoJj5Wqt9UY4rIieHNUObKugroDqWBJHepOMPmgmBO2pyK53Pqnhx4NIu/U
# ATHmIJh9yiq8ldOy2CpXiXJuMYIEDTCCBAkCAQEwgZMwfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAIhM8A1+9IPIaQAAQAAAiEwDQYJYIZIAWUDBAIB
# BQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQgzzyT+c4wWzg/YQtMpX2rew6bHiatkMy6fbGk27PWiBkwgfoGCyqGSIb3DQEJ
# EAIvMYHqMIHnMIHkMIG9BCAA7yEHnxVVGuAScvCGcsDAL5hkinVFahJsvQPvjwo9
# RDCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACITPA
# NfvSDyGkAAEAAAIhMCIEIC9tdFdvxtmm1/a0Wmzd+hrFYNbzL4yC0GJsTDu8U/r+
# MA0GCSqGSIb3DQEBCwUABIICADYr1ckXqQD9NJTMi0SGPZEFjQVn8igaztP95p1X
# ET1xlzREAGByikSWP88HEhycq922to9biJVrnkCXPSvW8iMa5mp76MbzhT4zfxza
# N3n3HCmHFLBuPeidkY2S8SQAEwhXloH6S4yfCz4J4G/Xf6eXi2Ap+t897ZH5DHxm
# GTM6ZXMtz5WlmHZNJt2MprvIuG6IMNBjdW4atAk0M0JvzRLK2js+bbAuK0nxl9Ar
# JPeBtXDkewIsOOpF7kD0rPcF9YeCIS+MYTH/dK4Gg9g3dAhI4RS2Ug3WIzO2CRSH
# 1xKJURbgutrQYs1VnbyBJMN+Sx3FbLga+yOWxtCxFh70yaaWaQ5Y5N8+pLpZpW00
# EXQVAZsPgIz8dTRYnhSQIB5p2R+W/E0NygzwNfWRyFLDNmniF0muJbc0ux0FHuqS
# DPm0D4YNDIlt5f5kJfxWhaq993+G6xwNK6e0NpbAKby8advItlkxLI9tHz4D15Qg
# +JIg/YVfTQLo0DuvC5auOThLBkHDWRMkXRs8VrSU5Kowqj7H9BHKYhQRLeX+tutZ
# /kBdGst4Tcqo99JPTLt/qQacXtSQp9zfErnMKFDcsOay9cQgUS7SI2A41kxBGX/0
# avahLoiomdXeGCorJok4kml/LrOb9yKwwQLWcrXVLJFFpbg1zoGxWFBIW8oMFxkI
# L6U7
# SIG # End signature block
