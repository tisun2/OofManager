@{

# Script module or binary module file associated with this manifest.
RootModule = 'Microsoft.PowerApps.Powershell.psm1'

# Version number of this module.
ModuleVersion = '1.0.45'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = 'e0c4f967-452b-43a0-a9f8-1f0ef9e06dd4'

# Author of this module
Author = 'Microsoft Common Data Service Team'

# Company or vendor of this module
CompanyName = 'Microsoft'

# Copyright statement for this module
Copyright = '© 2020 Microsoft Corporation. All rights reserved'

# Description of the functionality provided by this module
Description = 'PowerShell interface for Microsoft PowerApps and Flow features'

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
#RequiredModules = @("Microsoft.PowerApps.RestClientModule","Microsoft.PowerApps.AuthModule")

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
#NestedModules = @("Microsoft.PowerApps.AuthModule" , "Microsoft.PowerApps.RestClientModule") #,"Microsoft.PowerApps.Administration.PowerShell")

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @(
    'Get-PowerAppEnvironment', `
    'Get-PowerAppConnection', `
    'Remove-PowerAppConnection', `
    'Get-PowerAppConnectionRoleAssignment', `
    'Set-PowerAppConnectionRoleAssignment', `
    'Remove-PowerAppConnectionRoleAssignment', `
    'Get-PowerAppConnector', `
    'Remove-PowerAppConnector', `
    'Get-PowerAppConnectorRoleAssignment', `
    'Set-PowerAppConnectorRoleAssignment', `
    'Remove-PowerAppConnectorRoleAssignment', `
    'Get-PowerApp', `
    'Remove-PowerApp', `
    'Publish-PowerApp', `
    'Set-PowerAppPublishedAppSetting', `
    'Set-PowerAppDisplayName', `
    'Get-PowerAppVersion', `
    'Restore-PowerAppVersion', `
    'Get-PowerAppRoleAssignment', `
    'Set-PowerAppRoleAssignment', `
    'Remove-PowerAppRoleAssignment', `
    'Get-PowerAppsNotification', `
    'Set-PowerAppAsSolutionAware', `
    'Set-FlowAsSolutionAware', `
    'Get-FlowEnvironment', `
    'Get-Flow', `
    'Get-FlowOwnerRole', `
    'Set-FlowOwnerRole', `
    'Remove-FlowOwnerRole', `
    'Get-FlowRun', `
    'Enable-Flow', `
    'Disable-Flow', `
    'Remove-Flow', `
    'Get-FlowApprovalRequest', `
    'Get-FlowApproval', `
    'Approve-FlowApprovalRequest', `
	'Deny-FlowApprovalRequest', `
    'Add-FlowPowerAppContext', `
    'Remove-FlowPowerAppContext', `
    'Get-PowerVirtualAgentsDlpEnforcement', `
    'Set-PowerVirtualAgentsDlpEnforcement', `
    'Get-PowerPlatformRequestsConsumptionOfFlows', `
    'Regenerate-FlowAccessKey', `
	#from Rest and Auth Module Helpers
	'Select-CurrentEnvironment', `
	'Add-PowerAppsAccount', `
	'Remove-PowerAppsAccount',`
	'Test-PowerAppsAccount', `
	'Get-TenantDetailsFromGraph', `
	'Get-UsersOrGroupsFromGraph', `
	'Get-JwtToken', `
	'ReplaceMacro', `
	'InvokeApi'
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
ModuleList = @("Microsoft.PowerApps.PowerShell" )

# List of all files packaged with this module
# Note that Microsoft.IdentityModel.Clients.ActiveDirectory.dll and Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll are not included
# When included they are automatically loaded which can pull the files by name from uncontrolled locations.
FileList = @(
    "Microsoft.PowerApps.PowerShell.psm1", `
    "Microsoft.PowerApps.PowerShell.psd1", `
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
         IconUri = 'https://connectoricons-prod.azureedge.net/powerappsforappmakers/icon_1.0.1056.1255.png'

        # ReleaseNotes of this module
        ReleaseNotes = '

Current Release:
1.0.45
Fix a typo in Get-PowerAppConnector URL

1.0.9:
    Added usgov and usgovhigh support.'

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = 'PowerApp'

}

# SIG # Begin signature block
# MIIn/gYJKoZIhvcNAQcCoIIn7zCCJ+sCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDd2a640BhO/dpR
# yeQHWP8q2z+mlmcCAOXebF9XqMuG6KCCDXYwggX0MIID3KADAgECAhMzAAAEhV6Z
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGd4wghnaAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAASFXpnsDlkvzdcAAAAABIUwDQYJYIZIAWUDBAIB
# BQCggYIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJKoZIhvcNAQkEMSIE
# IFo6VOZpmzvAcqevMRRTSEOozomnD7GqwLJ+dWxqdCudMDQGCisGAQQBgjcCAQwx
# JjAkoBKAEABUAGUAcwB0AFMAaQBnAG6hDoAMaHR0cDovL3Rlc3QgMA0GCSqGSIb3
# DQEBAQUABIIBAE7crltEzYpMJ01k7ReDDJsAfQVFeLx/6toGQLLaoVmO66VlYcQp
# i/0GbskNcZzPYX9UAJ+Bt/zfqtztRfA2o7LU/Tr3h03YRBjjopxjWQzluHjAvkVp
# oHOkuPAYnLMpY+xpfgmn6SKe6aieBCVfIB28zdBhm/KDKwFCTN7pmcEFeEPz/tKA
# 8uKELrfCWnl+PxID0vVZ3DOgCBZci3w7pbgdzPFw5/j9Yb31FHN35USLDLV7D5SZ
# UJudKLVZ9pFM2xIReLlTOC94KxTv2EPLmOpBzmVc10yRE8wbEYveDL59ImPWLrGI
# G0w+I9Sw6Kz9kzTrx7jYHE1+t94FEc2sfAKhgheUMIIXkAYKKwYBBAGCNwMDATGC
# F4Awghd8BgkqhkiG9w0BBwKgghdtMIIXaQIBAzEPMA0GCWCGSAFlAwQCAQUAMIIB
# UgYLKoZIhvcNAQkQAQSgggFBBIIBPTCCATkCAQEGCisGAQQBhFkKAwEwMTANBglg
# hkgBZQMEAgEFAAQgMy3tBnFeNHeq8J1lDY3SySmo27kWwaRYqDZb8eA7irECBmjC
# xKxg4RgTMjAyNTA5MTUxNTU2MTkuMzcxWjAEgAIB9KCB0aSBzjCByzELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0
# IEFtZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkE5
# MzUtMDNFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
# aWNloIIR6jCCByAwggUIoAMCAQICEzMAAAIMuWTjNZzs9K4AAQAAAgwwDQYJKoZI
# hvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEm
# MCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjUwMTMw
# MTk0MzAwWhcNMjYwNDIyMTk0MzAwWjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0
# aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkE5MzUtMDNFMC1EOTQ3MSUw
# IwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAygFWJj3kbYGv1Oo86sqiw9DAKKJdt4HATefP
# mf17JGMDSbGfjvsFckGJfHT0ytfwQtsQInNZvib3JKo1WkN9iplEbBGaLYq0GODy
# lVvnl8Ebd6+rM4C7onOqqB5W16Mf5dBybYFEZMw11jJCphki+8/P3K6nL5mKr/Lf
# 7JQBeCfpqc2/bTBVZo8ClzjVXUcIPUN1mj2QQu1r6Iuz0SDdo4I0gZx2MgGUpbLS
# ja6WG+vhruqEhZEMxqUeufkDQ3ZD+Lnzn+D2zoN32+Lhj4yPBDypacDMGotZEMl/
# n4HIAqFfSfqPDGGAmVHrd5M4YcEc6oeizHg42lyz+9NUl14l3NmR87gx20v7GbSd
# +tu3FaQpVxCFL4Nsaa9Kz5SLR8LY6NT8DAqV2Kp2Cr1/GifJ2sE/VvBVLrsmTxtf
# OdvquI5FZXii+8fu3pfBE3oW3ZMHYQF8l4pmhM1nrTTUphvynxwKfXM8LC9byq+E
# YJ/qSCJGR7qJnX+XuPNSvsSFoSwj3ablfOxKhjiv424Tp2RKsHbwNAJTGi37Jgnp
# mZrqXo2mLhJNOf+nAlMYBeMwp5CXmHTAD/vWeJFYe7c0RbMP5WUpdg+xISAOip4+
# kX3x9pO2LUhkr/Ogkoc34l2s/curE7vEhqhejmy/3rvw5Ir8laAn1F1i44kibK0u
# tw9BBx0CAwEAAaOCAUkwggFFMB0GA1UdDgQWBBR1DkUh/7Af60P23g9JeVcUO9Oh
# iDAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSg
# UqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3Nv
# ZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEE
# YDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
# dDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB
# /wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEA2TD6+IFZsMH+BjAeWXx0q9+LoboO
# ss7uB1E/iVjGas/boS2QaF+Qj43Sic8AFb2KDbi5ktPvZQOUu+K7yqnf7vb6fPFR
# pOlO4DHHmrXaqSpW1UXZ9mX6zHKSOMznOgbbmK8yVeHBLNWJl/ebogMWhA9+MNNg
# Z37j2VwNHnbAwW3eIsRVPF/9SdA3yFJNWBWDzq5sJiNpNeruk3CjtGKUZcE3Qqvb
# ztHhCBEdUi5kDQc1/YdnHAr7YHpDmgaCEN2UWovA7NX/sHCgj8w+Kg198TYLyxYi
# qAOmUhvUv8jqxmokhiHg8uTfVULqkzY68rgM473+VvAEKd9YVdRm1AzpG1HXfs5C
# Vil+BZs3njedhBG8pKFnCeVfTOAzxjecaRal8vWjtPnUdFFGFrqni4Q8kZ1XmXEx
# LtMYJqPqUB2rhVQErFTkTKfExfHaXrHfrapJEPFTbyNtKDn503y/u2YFDH+6jVdJ
# ZdFqOZ5a9Qib2tW35Nh3OQWNTPbHd25QZHs8ryT5+I9G3zjqwmE8GLDbI4kZf1lt
# fDTqYsKnIsBZVDarVgkTMwva/OGGlDEPNgcsJOPHeLgaJ+WQPKV10u48CU4yY+VE
# nkZfb40/fDw2cghTtnhUjhXQ3X+lgaP1mVANoRmdKvie49eNH21wnzlCJtI9tx2g
# FdHJA0v55gv6BdYwggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0G
# CSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3Jp
# dHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9
# uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZr
# BxTzxXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk
# 2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxR
# nOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uD
# RedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGa
# RnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fz
# pk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG
# 4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGU
# lNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLE
# hReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0w
# ggHZMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+
# gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNV
# HSAEVTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0P
# BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9
# lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3Js
# MFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJ
# KoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEG
# k5c9MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2
# LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7nd
# n/OOPcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSF
# QrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy8
# 7JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8
# x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2f
# pCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz
# /gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQ
# KBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAx
# M328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGby
# oYIDTTCCAjUCAQEwgfmhgdGkgc4wgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpBOTM1LTAzRTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIa
# AxUA77vIZIRDLeWfC3Xn5bO89S1VPKaggYMwgYCkfjB8MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1T
# dGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQsFAAIFAOxyiMEwIhgPMjAyNTA5MTUx
# MjQ0NDlaGA8yMDI1MDkxNjEyNDQ0OVowdDA6BgorBgEEAYRZCgQBMSwwKjAKAgUA
# 7HKIwQIBADAHAgEAAgIKATAHAgEAAgITFzAKAgUA7HPaQQIBADA2BgorBgEEAYRZ
# CgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0G
# CSqGSIb3DQEBCwUAA4IBAQB6hAbSYlan/BKNm3ZOlOv64my0WiH6Bk46Z6KNQX5f
# ceAtbYHlYd5urbwY++uigICtYOTcJekMxFUv2RJBIp/VCyL/H8ThYOfwDyEvtUSw
# SHUsuSqborb1yC8K5/rRxSjrYn/GYJD7vhGNcx4n1H3P64qfV8XAcQRJzhI9RX+C
# H0dwEN6nYsbe4ZA3ZgYpF6Atgc4bQI6DJyvA35jwPeEU9b+s1wWlXye9TjgglxVe
# +32neDS1rBcOyAyZxwS7ELfeU7JY22TsHhnaGC3P6a6yhfnNSWMzkgaPzzbeRrRm
# sBnq11tyDsrUE4bfcCvEp4GzRF014hslihv3YmjM+EwxMYIEDTCCBAkCAQEwgZMw
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIMuWTjNZzs9K4AAQAA
# AgwwDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRAB
# BDAvBgkqhkiG9w0BCQQxIgQgFBhb+J8yG6jDCZN68IoycV4KCr+P1EmttMgcjY30
# kogwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCDVKNe3BTGTeOjCOTXyAIPV
# MeXDucTPYp63ua4rjmfCLTCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAACDLlk4zWc7PSuAAEAAAIMMCIEIIT7zyP8dj/LhBw6nrmKaZlQ
# NE14LsMW8oh6H3L/ceIoMA0GCSqGSIb3DQEBCwUABIICAF1r3OiN/aPeCnl1HNJu
# E/bI6FnHKjHUIAOs2tInot9Vv2Dh4Vm1MTW7xrewyeqaBc505OXUzzjCCZbSS7IA
# FSrm2eKAkwCUNc1JZ3nFTY2l0RAFNulB3YrBSThcaWYZE+6i31vLys72W8SPfV2A
# vdDVI3qdxDgXLryq/X3QULq5uoDCdsm7QrIWksMFcufVAWJzQJqEB2gWLPDcOs+A
# D3fYeQhgQLemKSssgqXA0zniLSDdetG15iYrcgU0ifraEmDs3bX9/d/v7g/96Nax
# /cgE4UUnM3Hx1yGoYx3o6GANvadmlCDOsGUiugy0i+YXTBQd9+UYmN4IYkoQGcXe
# dfSdvw6dlSkR/r5FPy/WsrCCcVjq9K2mkJQrGiVFwbOXDhs7MGEOUnXNP6T7wfjC
# LV6rHcUVoThBEqCw6LyFCfbDifn1MoJs9fM9jY10PmXB6c90fqPPBfR55XFU7EqI
# N2BIFxPDHobEdrJWIz8V9JTHqwLsDmrNkXfB55lPWGNg+MKRfVMan3vinbRW9n/+
# PW+XWVFbGmXol+T5/q1IVNUfifFpXeckn6usQo49dKEnJ2hQpYDc2FGahsgiw1jy
# UFZY3cDEFTkCfdXuNg9ch2S1KnY+ZbzCQ5lMHMGSQJ7HlwqnLbTL9Sodneq8Xxn3
# mNkKeejO6n9sBifB5Tm5sAkz
# SIG # End signature block
