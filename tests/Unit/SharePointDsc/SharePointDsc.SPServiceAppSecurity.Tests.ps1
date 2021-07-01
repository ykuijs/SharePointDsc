[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
param
(
    [Parameter()]
    [string]
    $SharePointCmdletModule = (Join-Path -Path $PSScriptRoot `
            -ChildPath "..\Stubs\SharePoint\15.0.4805.1000\Microsoft.SharePoint.PowerShell.psm1" `
            -Resolve)
)

$script:DSCModuleName = 'SharePointDsc'
$script:DSCResourceName = 'SPServiceAppSecurity'
$script:DSCResourceFullName = 'MSFT_' + $script:DSCResourceName

function Invoke-TestSetup
{
    try
    {
        Import-Module -Name DscResource.Test -Force

        Import-Module -Name (Join-Path -Path $PSScriptRoot `
                -ChildPath "..\UnitTestHelper.psm1" `
                -Resolve)

        $Global:SPDscHelper = New-SPDscUnitTestHelper -SharePointStubModule $SharePointCmdletModule `
            -DscResource $script:DSCResourceName
    }
    catch [System.IO.FileNotFoundException]
    {
        throw 'DscResource.Test module dependency not found. Please run ".\build.ps1 -Tasks build" first.'
    }

    $script:testEnvironment = Initialize-TestEnvironment `
        -DSCModuleName $script:DSCModuleName `
        -DSCResourceName $script:DSCResourceFullName `
        -ResourceType 'Mof' `
        -TestType 'Unit'
}

function Invoke-TestCleanup
{
    Restore-TestEnvironment -TestEnvironment $script:testEnvironment
}

Invoke-TestSetup

try
{
    InModuleScope -ModuleName $script:DSCResourceFullName -ScriptBlock {
        Describe -Name $Global:SPDscHelper.DescribeHeader -Fixture {
            BeforeAll {
                Invoke-Command -Scriptblock $Global:SPDscHelper.InitializeScript -NoNewScope

                # Initialize tests

                # Mocks for all contexts
                Mock -CommandName Test-SPDscIsADUser -MockWith { return $true }

                Mock Grant-SPObjectSecurity -MockWith { }
                Mock Revoke-SPObjectSecurity -MockWith { }
                Mock -CommandName Set-SPServiceApplicationSecurity -MockWith { }

                Mock -CommandName New-SPClaimsPrincipal -MockWith {
                    return @{
                        Value = $Identity -replace "i:0#.w\|"
                    }
                } -ParameterFilter { $IdentityType -eq "EncodedClaim" }

                Mock -CommandName New-SPClaimsPrincipal -MockWith {
                    $Global:SPDscClaimsPrincipalUser = $Identity
                    return (
                        New-Object -TypeName "Object" | Add-Member -MemberType ScriptMethod `
                            -Name ToEncodedString `
                            -Value {
                            return "i:0#.w|$($Global:SPDscClaimsPrincipalUser)"
                        } -PassThru
                    )
                } -ParameterFilter { $IdentityType -eq "WindowsSamAccountName" }

                Mock -CommandName Get-SPFarm -MockWith {
                    return @{
                        Id = [Guid]"02a0cea2-d4e0-4e4e-ba2e-e532a433cfef"
                    }
                }

                function Add-SPDscEvent
                {
                    param (
                        [Parameter(Mandatory = $true)]
                        [System.String]
                        $Message,

                        [Parameter(Mandatory = $true)]
                        [System.String]
                        $Source,

                        [Parameter()]
                        [ValidateSet('Error', 'Information', 'FailureAudit', 'SuccessAudit', 'Warning')]
                        [System.String]
                        $EntryType,

                        [Parameter()]
                        [System.UInt32]
                        $EventID
                    )
                }
            }

            # Test contexts
            Context -Name "The service app that security should be applied to does not exist" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                }),
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user2"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return $null
                    }
                }

                It "Should return empty members list from the get method" {
                    (Get-TargetResource @testParams).Members | Should -BeNullOrEmpty
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should throw an exception in the set method" {
                    { Set-TargetResource @testParams } | Should -Throw
                }
            }

            Context -Name "None of the required members properties are provided" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                    }
                }

                It "Should throw an exception from the test method" {
                    { Test-TargetResource @testParams } | Should -Throw
                }

                It "Should throw an exception from the set method" {
                    { Set-TargetResource @testParams } | Should -Throw
                }
            }

            Context -Name "All of the members properties are provided" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        Members          = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                })
                        )
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                })
                        )
                        MembersToExclude = @("CONTOSO\user2")
                    }
                }

                It "Should throw an exception from the test method" {
                    { Test-TargetResource @testParams } | Should -Throw
                }

                It "Should throw an exception from the set method" {
                    { Set-TargetResource @testParams } | Should -Throw
                }
            }

            Context -Name "A specified access level does not match the allowed list of (localized) access levels (Members)" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @(
                            (New-CimInstance -ClassName "MSFT_SPSearchServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Read"
                                }),
                            (New-CimInstance -ClassName "MSFT_SPSearchServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user2"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        return @{
                            NamedAccessRights = @(
                                @{
                                    Name = "Full Control"
                                }
                            )
                        }
                    }
                }

                It "Should return en empty ServiceAppName from the get method" {
                    (Get-TargetResource @testParams).ServiceAppName | Should -Be ""
                }

                It "Should return False from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should throw an exception in the set method" {
                    { Set-TargetResource @testParams } | Should -Throw "Unknown AccessLevel is used"
                }
            }

            Context -Name "A specified access level does not match the allowed list of (localized) access levels (MembersToInclude)" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPSearchServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Read"
                                }),
                            (New-CimInstance -ClassName "MSFT_SPSearchServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user2"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        return @{
                            NamedAccessRights = @(
                                @{
                                    Name = "Full Control"
                                }
                            )
                        }
                    }
                }

                It "Should return en empty ServiceAppName from the get method" {
                    (Get-TargetResource @testParams).ServiceAppName | Should -Be ""
                }

                It "Should return False from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should throw an exception in the set method" {
                    { Set-TargetResource @testParams } | Should -Throw "Unknown AccessLevel is used"
                }
            }

            Context -Name "The service app exists and a fixed members list is provided that does not match the current settings" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @(
                            (New-CimInstance -ClassName "MSFT_SPSearchServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                }),
                            (New-CimInstance -ClassName "MSFT_SPSearchServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user2"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        if ($Global:SPDscRunCount -in 0, 1, 3, 4)
                        {
                            $Global:SPDscRunCount++
                            return @{
                                NamedAccessRights = @(
                                    @{
                                        Name = "Full Control"
                                    }
                                )
                            }
                        }
                        else
                        {
                            $Global:SPDscRunCount++
                            return @{
                                AccessRules = @(
                                    @{
                                        Name          = "CONTOSO\user1"
                                        AllowedRights = "Read"
                                    }
                                )
                            }
                        }
                    }
                }

                It "Should return a list of current members from the get method" {
                    $Global:SPDscRunCount = 0
                    (Get-TargetResource @testParams).Members | Should -Not -BeNullOrEmpty
                }

                It "Should return false from the test method" {
                    $Global:SPDscRunCount = 0
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    $Global:SPDscRunCount = 0
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity -ParameterFilter { $Replace -eq $true }
                    Assert-MockCalled Set-SPServiceApplicationSecurity
                }
            }

            Context -Name "The service app exists and a fixed members list is provided that does match the current settings" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                }),
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user2"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        $security = @{
                            NamedAccessRights = @(
                                @{
                                    Name   = "Full Control"
                                    Rights = @{
                                        RightsFlags = 0xff
                                    }
                                }
                            )
                            AccessRules       = @(
                                @{
                                    Name                = "CONTOSO\user1"
                                    AllowedObjectRights =
                                    @{
                                        RightsFlags = 0xff
                                    }
                                }
                                @{
                                    Name                = "CONTOSO\user2"
                                    AllowedObjectRights =
                                    @{
                                        RightsFlags = 0xff
                                    }
                                }
                            )
                        }

                        $security.NamedAccessRights | ForEach-Object { $_.Rights } | Add-Member -MemberType ScriptMethod `
                            -Name IsSubsetOf `
                            -Value {
                            param($objectRights)
                            return ($objectRights.RightsFlags -band $this.RightsFlags) -eq $this.RightsFlags
                        }

                        return $security
                    }
                }

                It "Should return a list of current members from the get method" {
                    (Get-TargetResource @testParams).Members | Should -Not -BeNullOrEmpty
                }

                It "Should return true from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "The service app exists and a specific list of members to add and remove is provided, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                })
                        )
                        MembersToExclude = @("CONTOSO\user2")
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        if ($Global:SPDscRunCount -in 0, 1, 3, 4)
                        {
                            $Global:SPDscRunCount++
                            return @{
                                NamedAccessRights = @(
                                    @{
                                        Name = "Full Control"
                                    }
                                )
                            }
                        }
                        else
                        {
                            $Global:SPDscRunCount++
                            return @{
                                AccessRules = @(
                                    @{
                                        Name          = "CONTOSO\user2"
                                        AllowedRights = "FullControl"
                                    }
                                )
                            }
                        }
                    }
                }

                It "Should return a list of current members from the get method" {
                    $Global:SPDscRunCount = 0
                    (Get-TargetResource @testParams).Members | Should -Not -BeNullOrEmpty
                }

                It "Should return false from the test method" {
                    $Global:SPDscRunCount = 0
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    $Global:SPDscRunCount = 0
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity
                    Assert-MockCalled Revoke-SPObjectSecurity
                    Assert-MockCalled Set-SPServiceApplicationSecurity
                }
            }

            Context -Name "The service app exists and a specific list of members to remove is provided, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToExclude = @("CONTOSO\user2")
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        return @{
                            AccessRules = @(
                                @{
                                    Name          = "CONTOSO\user2"
                                    AllowedRights = "FullControl"
                                }
                            )
                        }
                    }
                }

                It "Should return a list of current members from the get method" {
                    (Get-TargetResource @testParams).Members | Should -Not -BeNullOrEmpty
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Revoke-SPObjectSecurity
                    Assert-MockCalled Set-SPServiceApplicationSecurity
                }
            }

            Context -Name "The service app exists and a specific list of members to add and remove is provided, which does match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                })
                        )
                        MembersToExclude = @("CONTOSO\user2")
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity {
                        $security = @{
                            NamedAccessRights = @(
                                @{
                                    Name   = "Full Control"
                                    Rights = @{
                                        RightsFlags = 0xff
                                    }
                                }
                            )
                            AccessRules       = @(
                                @{
                                    Name                = "CONTOSO\user1"
                                    AllowedObjectRights =
                                    @{
                                        RightsFlags = 0xff
                                    }
                                }
                            )
                        }

                        $security.NamedAccessRights | ForEach-Object { $_.Rights } | Add-Member -MemberType ScriptMethod `
                            -Name IsSubsetOf `
                            -Value {
                            param($objectRights)
                            return ($objectRights.RightsFlags -band $this.RightsFlags) -eq $this.RightsFlags
                        }

                        return $security
                    }
                }

                It "Should return a list of current members from the get method" {
                    (Get-TargetResource @testParams).Members | Should -Not -BeNullOrEmpty
                }

                It "Should return true from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "The service app exists and a specific list of members to add is provided with different access level, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Read"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        if ($Global:SPDscRunCount -in 0, 1)
                        {
                            $Global:SPDscRunCount++
                            return @{
                                NamedAccessRights = @(
                                    @{
                                        Name = "Full Control"
                                    },
                                    @{
                                        Name = "Read"
                                    }
                                )
                            }
                        }
                        else
                        {
                            $Global:SPDscRunCount++
                            return @{
                                AccessRules = @(
                                    @{
                                        Name          = "CONTOSO\user1"
                                        AllowedRights = "FullControl"
                                    }
                                )
                            }
                        }
                    }
                }

                It "Should return false from the test method" {
                    $Global:SPDscRunCount = 0
                    Test-TargetResource @testParams | Should -Be $false
                }
            }

            Context -Name "The service app exists and a specific list of members is provided with different access level, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Read"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        if ($Global:SPDscRunCount -in 0, 1, 3, 4)
                        {
                            $Global:SPDscRunCount++
                            return @{
                                NamedAccessRights = @(
                                    @{
                                        Name = "Read"
                                    }
                                )
                            }
                        }
                        else
                        {
                            $Global:SPDscRunCount++
                            return @{
                                AccessRules = @(
                                    @{
                                        Name          = "CONTOSO\user1"
                                        AllowedRights = "FullControl"
                                    }
                                )
                            }
                        }
                    }
                }

                It "Should return false from the test method" {
                    $Global:SPDscRunCount = 0
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    $Global:SPDscRunCount = 0
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity -Times 1 -ParameterFilter { $Replace -eq $true }
                    Assert-MockCalled Set-SPServiceApplicationSecurity -Times 1
                }
            }

            Context -Name "The service app exists and a specific list of members to add is provided with different access level, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Read"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        if ($Global:SPDscRunCount -in 0, 1, 3, 4)
                        {
                            $Global:SPDscRunCount++
                            return @{
                                NamedAccessRights = @(
                                    @{
                                        Name = "Read"
                                    }
                                )
                            }
                        }
                        else
                        {
                            $Global:SPDscRunCount++
                            return @{
                                AccessRules = @(
                                    @{
                                        Name          = "CONTOSO\user1"
                                        AllowedRights = "FullControl"
                                    }
                                )
                            }
                        }
                    }
                }

                It "Should return false from the test method" {
                    $Global:SPDscRunCount = 0
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    $Global:SPDscRunCount = 0
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity -Times 1 -ParameterFilter { $Replace -eq $true }
                    Assert-MockCalled Set-SPServiceApplicationSecurity -Times 1
                }
            }

            Context -Name "The service app exists and an empty list of members is provided, which matches the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @()
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        return @{
                            AccessRules = @()
                        }
                    }
                }

                It "Should return true from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
                It "Should call the update cmdlet from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity -Times 0
                    Assert-MockCalled Set-SPServiceApplicationSecurity -Times 1
                }
            }

            Context -Name "The service app exists and an empty list of members is provided, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @()
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        return @{
                            AccessRules = @(
                                @{
                                    Name          = "CONTOSO\user1"
                                    AllowedRights = "FullControl"
                                }
                            )
                        }
                    }
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }
                It "Should call the update cmdlet from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity -Times 0
                    Assert-MockCalled Revoke-SPObjectSecurity -Times 1
                    Assert-MockCalled Set-SPServiceApplicationSecurity -Times 1
                }
            }

            Context -Name "The service app exists and a specific list of members to add and remove is provided, which does match the desired state and includes a claims based group" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Full Control"
                                })
                        )
                        MembersToExclude = @("CONTOSO\user2")
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }
                    Mock -CommandName Get-SPServiceApplicationSecurity {
                        $security = @{
                            NamedAccessRights = @(
                                @{
                                    Name   = "Full Control"
                                    Rights = @{
                                        RightsFlags = 0xff
                                    }

                                }
                            )
                            AccessRules       = @(
                                @{
                                    Name                = "i:0#.w|s-1-5-21-2753725054-2932589700-2007370523-2138"
                                    AllowedObjectRights =
                                    @{
                                        RightsFlags = 0xff
                                    }
                                }
                            )
                        }

                        $security.NamedAccessRights | ForEach-Object { $_.Rights } | Add-Member -MemberType ScriptMethod `
                            -Name IsSubsetOf `
                            -Value {
                            param($objectRights)
                            return ($objectRights.RightsFlags -band $this.RightsFlags) -eq $this.RightsFlags
                        }

                        return $security
                    }

                    Mock Resolve-SPDscSecurityIdentifier {
                        return "CONTOSO\user1"
                    }
                }

                It "Should return a list of current members from the get method" {
                    (Get-TargetResource @testParams).Members | Should -Not -BeNullOrEmpty
                }

                It "Should return true from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "The service app exists and the local farm token is provided in the members list that match the current settings" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "{LocalFarm}"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity {
                        $security = @{
                            NamedAccessRights = @(
                                @{
                                    Name   = "Full Control"
                                    Rights = @{
                                        RightsFlags = 0xff
                                    }

                                }
                            )
                            AccessRules       = @(
                                @{
                                    Name                = "c:0%.c|system|02a0cea2-d4e0-4e4e-ba2e-e532a433cfef"
                                    AllowedObjectRights =
                                    @{
                                        RightsFlags = 0xff
                                    }
                                }
                            )
                        }

                        $security.NamedAccessRights | ForEach-Object { $_.Rights } | Add-Member -MemberType ScriptMethod `
                            -Name IsSubsetOf `
                            -Value {
                            param($objectRights)
                            return ($objectRights.RightsFlags -band $this.RightsFlags) -eq $this.RightsFlags
                        }

                        return $security
                    }
                }

                It "Should return local farm token in the list of current members from the get method" {
                    $members = (Get-TargetResource @testParams).Members
                    $members[0].Username | Should -Be "{LocalFarm}"
                }

                It "Should return true from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "The service app exists and the local farm token is provided in the members list that does not match the current settings" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "{LocalFarm}"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        if ($Global:SPDscRunCount -in 0, 1, 3, 4)
                        {
                            $Global:SPDscRunCount++
                            return @{
                                NamedAccessRights = @(
                                    @{
                                        Name = "Full Control"
                                    }
                                )
                            }
                        }
                        else
                        {
                            $Global:SPDscRunCount++
                            return @{
                                AccessRules = @(
                                )
                            }
                        }
                    }
                }

                It "Should return false from the test method" {
                    $Global:SPDscRunCount = 0
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    $Global:SPDscRunCount = 0
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity
                    Assert-MockCalled Revoke-SPObjectSecurity -Times 0
                    Assert-MockCalled Set-SPServiceApplicationSecurity
                }
            }

            Context -Name "The service app exists and local farm token is included in the specific list of members to add, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "{LocalFarm}"
                                    AccessLevels = "Full Control"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        if ($Global:SPDscRunCount -in 0, 1, 3, 4)
                        {
                            $Global:SPDscRunCount++
                            return @{
                                NamedAccessRights = @(
                                    @{
                                        Name = "Full Control"
                                    }
                                )
                            }
                        }
                        else
                        {
                            $Global:SPDscRunCount++
                            return @{
                                AccessRules = @(
                                    @{
                                        Name          = "CONTOSO\user2"
                                        AllowedRights = "FullControl"
                                    }
                                )
                            }
                        }
                    }
                }

                It "Should return false from the test method" {
                    $Global:SPDscRunCount = 0
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    $Global:SPDscRunCount = 0
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity
                    Assert-MockCalled Revoke-SPObjectSecurity -Times 0
                    Assert-MockCalled Set-SPServiceApplicationSecurity
                }
            }

            Context -Name "The service app exists and local farm token is included in the specific list of members to remove, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToExclude = @("{LocalFarm}")
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        return @{
                            AccessRules = @(
                                @{
                                    Name          = "c:0%.c|system|02a0cea2-d4e0-4e4e-ba2e-e532a433cfef"
                                    AllowedRights = "FullControl"
                                }
                            )
                        }
                    }
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity -Times 0
                    Assert-MockCalled Revoke-SPObjectSecurity
                    Assert-MockCalled Set-SPServiceApplicationSecurity
                }
            }

            Context -Name "The service app exists and an empty list of members are specified, which does not match the desired state" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName = "Example Service App"
                        SecurityType   = "SharingPermissions"
                        Members        = @()
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity -MockWith {
                        return @{
                            AccessRules = @(
                                @{
                                    Name          = "c:0%.c|system|02a0cea2-d4e0-4e4e-ba2e-e532a433cfef"
                                    AllowedRights = "FullControl"
                                }
                            )
                        }
                    }
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update cmdlet from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Grant-SPObjectSecurity -Times 0
                    Assert-MockCalled Revoke-SPObjectSecurity
                    Assert-MockCalled Set-SPServiceApplicationSecurity
                }
            }

            Context -Name "Access level that includes other access levels is specified when multiple named access rights exists" -Fixture {
                BeforeAll {
                    $testParams = @{
                        ServiceAppName   = "Example Service App"
                        SecurityType     = "SharingPermissions"
                        MembersToInclude = @(
                            (New-CimInstance -ClassName "MSFT_SPServiceAppSecurityEntry" `
                                    -ClientOnly `
                                    -Property @{
                                    Username     = "CONTOSO\user1"
                                    AccessLevels = "Contribute"
                                })
                        )
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return @{
                            Name = $testParams.ServiceAppName
                        }
                    }

                    Mock -CommandName Get-SPServiceApplicationSecurity {
                        $security = @{
                            NamedAccessRights = @(
                                @{
                                    Name   = "Full Control"
                                    Rights = @{
                                        RightsFlags = 0xff
                                    }
                                }
                                @{
                                    Name   = "Contribute"
                                    Rights = @{
                                        RightsFlags = 0x0f
                                    }
                                }
                                @{
                                    Name   = "Read"
                                    Rights = @{
                                        RightsFlags = 0x01
                                    }
                                }
                            )
                            AccessRules       = @(
                                @{
                                    Name                = "CONTOSO\user1"
                                    AllowedObjectRights =
                                    @{
                                        RightsFlags = 0x0f
                                    }
                                }
                            )
                        }

                        $security.NamedAccessRights | ForEach-Object { $_.Rights } | Add-Member -MemberType ScriptMethod `
                            -Name IsSubsetOf `
                            -Value {
                            param($objectRights)
                            return ($objectRights.RightsFlags -band $this.RightsFlags) -eq $this.RightsFlags
                        }

                        return $security
                    }
                }

                It "Should return list of all named access rights included for user" {
                    $result = Get-TargetResource @testParams
                    $result.Members | Should -Not -BeNullOrEmpty
                    $result.Members[0].AccessLevels | Should -Be @("Contribute", "Read")
                }

                It "Should return true from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "Running ReverseDsc Export" -Fixture {
                BeforeAll {
                    Import-Module (Join-Path -Path (Split-Path -Path (Get-Module SharePointDsc -ListAvailable).Path -Parent) -ChildPath "Modules\SharePointDSC.Reverse\SharePointDSC.Reverse.psm1")

                    Mock -CommandName Write-Host -MockWith { }

                    Mock -CommandName Get-TargetResource -MockWith {
                        return @{
                            ServiceAppName = "User Profile Service Application"
                            SecurityType   = "Administrators"
                            Members        = @(
                                @{
                                    Username     = "Domain\Account"
                                    AccessLevels = "Full Control"
                                }
                            )
                        }
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            DisplayName = "User Profile Service Application"
                            Name        = "User Profile Service Application"
                        }
                        $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod `
                            -Name GetType `
                            -Value {
                            return @{
                                Name = "User Profile Service Application"
                            }
                        } -PassThru -Force
                        return $spServiceApp
                    }

                    if ($null -eq (Get-Variable -Name 'spFarmAccount' -ErrorAction SilentlyContinue))
                    {
                        $mockPassword = ConvertTo-SecureString -String "password" -AsPlainText -Force
                        $Global:spFarmAccount = New-Object -TypeName System.Management.Automation.PSCredential ("contoso\spfarm", $mockPassword)
                    }

                    $result = @'
        \$members = \@\(\);
        \$members \+= MSFT_SPServiceAppSecurityEntry
        {
            Username     = "Domain\\Account";
            AccessLevels = \@\("Full Control"\);
        };
        SPServiceAppSecurity [0-9A-Fa-f]{8}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{12}
        {
            Members              = "\$members";
            PsDscRunAsCredential = \$Credsspfarm;
            SecurityType         = "Administrators";
            ServiceAppName       = "User Profile Service Application";
        }
        \$members = \@\(\);
        \$members \+= MSFT_SPServiceAppSecurityEntry
        {
            Username     = "Domain\\Account";
            AccessLevels = \@\("Full Control"\);
        };
        SPServiceAppSecurity [0-9A-Fa-f]{8}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{12}
        {
            Members              = "\$members";
            PsDscRunAsCredential = \$Credsspfarm;
            SecurityType         = "Administrators";
            ServiceAppName       = "User Profile Service Application";
        }

'@
                }

                It "Should return valid DSC block from the Export method" {
                    Export-TargetResource | Should -Match $result
                }
            }
        }
    }
}
finally
{
    Invoke-TestCleanup
}
