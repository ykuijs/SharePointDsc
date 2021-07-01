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
$script:DSCResourceName = 'SPUsageApplication'
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
                $getTypeFullName = "Microsoft.SharePoint.Administration.SPUsageApplication"
                $getTypeFullNameProxy = "Microsoft.SharePoint.Administration.SPUsageApplicationProxy"

                $mockPassword = ConvertTo-SecureString -String "password" -AsPlainText -Force
                $mockCredential = New-Object -TypeName System.Management.Automation.PSCredential `
                    -ArgumentList @("DOMAIN\username", $mockPassword)

                # Mocks for all contexts
                Mock -CommandName New-SPUsageApplication -MockWith { }
                Mock -CommandName Set-SPUsageService -MockWith { }
                Mock -CommandName Get-SPUsageService -MockWith {
                    return @{
                        UsageLogCutTime     = $testParams.UsageLogCutTime
                        UsageLogDir         = $testParams.UsageLogLocation
                        UsageLogMaxFileSize = ($testParams.UsageLogMaxFileSizeKB * 1024)
                        UsageLogMaxSpaceGB  = $testParams.UsageLogMaxSpaceGB
                    }
                }
                Mock -CommandName Remove-SPServiceApplication -MockWith { }
                Mock -CommandName Get-SPServiceApplicationProxy -MockWith {
                    return (New-Object -TypeName "Object" |
                            Add-Member -MemberType ScriptMethod `
                                -Name Provision `
                                -Value { } `
                                -PassThru |
                                Add-Member -NotePropertyName Status `
                                    -NotePropertyValue "Online" `
                                    -PassThru |
                                    Add-Member -NotePropertyName TypeName `
                                        -NotePropertyValue "Usage and Health Data Collection Proxy" `
                                        -PassThru)
                }
            }

            # Test contexts
            Context -Name "When no service applications exist in the current farm" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name                   = "Usage Service App"
                        UsageLogCutTime        = 60
                        UsageLogLocation       = "L:\UsageLogs"
                        UsageLogMaxFileSizeKB  = 1024
                        UsageLogMaxSpaceGB     = 10
                        DatabaseName           = "SP_Usage"
                        DatabaseServer         = "sql.test.domain"
                        FailoverDatabaseServer = "anothersql.test.domain"
                        Ensure                 = "Present"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith { return $null }
                }

                It "Should return null from the Get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Absent"
                }

                It "Should return false when the Test method is called" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should create a new service application in the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled New-SPUsageApplication
                }

                It "Should create a new service application with custom database credentials" {
                    $testParams.Add("DatabaseCredentials", $mockCredential)
                    Set-TargetResource @testParams
                    Assert-MockCalled New-SPUsageApplication
                }
            }

            Context -Name "When service applications exist in the current farm but not the specific usage service app" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name                   = "Usage Service App"
                        UsageLogCutTime        = 60
                        UsageLogLocation       = "L:\UsageLogs"
                        UsageLogMaxFileSizeKB  = 1024
                        UsageLogMaxSpaceGB     = 10
                        DatabaseName           = "SP_Usage"
                        DatabaseServer         = "sql.test.domain"
                        FailoverDatabaseServer = "anothersql.test.domain"
                        Ensure                 = "Present"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            DisplayName = $testParams.Name
                        }
                        $spServiceApp | Add-Member -MemberType ScriptMethod `
                            -Name GetType `
                            -Value {
                            return @{
                                FullName = "Microsoft.Office.UnKnownWebServiceApplication"
                            }
                        } -PassThru -Force
                        return $spServiceApp
                    }
                }

                It "Should return absent from the Get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Absent"
                }

                It "Should return false when the Test method is called" {
                    Test-TargetResource @testParams | Should -Be $false
                }
            }

            Context -Name "When a service application exists and is configured correctly" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name                   = "Usage Service App"
                        UsageLogCutTime        = 60
                        UsageLogLocation       = "L:\UsageLogs"
                        UsageLogMaxFileSizeKB  = 1024
                        UsageLogMaxSpaceGB     = 10
                        DatabaseName           = "SP_Usage"
                        DatabaseServer         = "sql.test.domain"
                        FailoverDatabaseServer = "anothersql.test.domain"
                        Ensure                 = "Present"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            TypeName      = "Usage and Health Data Collection Service Application"
                            DisplayName   = $testParams.Name
                            Name          = $testParams.Name
                            UsageDatabase = @{
                                Name                 = $testParams.DatabaseName
                                NormalizedDataSource = $testParams.DatabaseServer
                            }
                        }
                        $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                            return @{ FullName = $getTypeFullName }
                        } -PassThru -Force
                        return $spServiceApp
                    }
                }

                It "Should return values from the get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Present"
                }

                It "Should return true when the Test method is called" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "When a service application exists and log path are not configured correctly" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name                   = "Usage Service App"
                        UsageLogCutTime        = 60
                        UsageLogLocation       = "L:\UsageLogs"
                        UsageLogMaxFileSizeKB  = 1024
                        UsageLogMaxSpaceGB     = 10
                        DatabaseName           = "SP_Usage"
                        DatabaseServer         = "sql.test.domain"
                        FailoverDatabaseServer = "anothersql.test.domain"
                        Ensure                 = "Present"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            TypeName      = "Usage and Health Data Collection Service Application"
                            DisplayName   = $testParams.Name
                            Name          = $testParams.Name
                            UsageDatabase = @{
                                Name                 = $testParams.DatabaseName
                                NormalizedDataSource = $testParams.DatabaseServer
                            }
                        }
                        $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                            return @{ FullName = $getTypeFullName }
                        } -PassThru -Force
                        return $spServiceApp
                    }

                    Mock -CommandName Get-SPUsageService -MockWith {
                        return @{
                            UsageLogCutTime     = $testParams.UsageLogCutTime
                            UsageLogDir         = "C:\Wrong\Location"
                            UsageLogMaxFileSize = ($testParams.UsageLogMaxFileSizeKB * 1024)
                            UsageLogMaxSpaceGB  = $testParams.UsageLogMaxSpaceGB
                        }
                    }
                }

                It "Should return false when the Test method is called" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update service app cmdlet from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Set-SPUsageService
                }
            }

            Context -Name "When a service application exists and log size is not configured correctly" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name                   = "Usage Service App"
                        UsageLogCutTime        = 60
                        UsageLogLocation       = "L:\UsageLogs"
                        UsageLogMaxFileSizeKB  = 1024
                        UsageLogMaxSpaceGB     = 10
                        DatabaseName           = "SP_Usage"
                        DatabaseServer         = "sql.test.domain"
                        FailoverDatabaseServer = "anothersql.test.domain"
                        Ensure                 = "Present"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            TypeName      = "Usage and Health Data Collection Service Application"
                            DisplayName   = $testParams.Name
                            Name          = $testParams.Name
                            UsageDatabase = @{
                                Name                 = $testParams.DatabaseName
                                NormalizedDataSource = $testParams.DatabaseServer
                            }
                        }
                        $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                            return @{ FullName = $getTypeFullName }
                        } -PassThru -Force
                        return $spServiceApp
                    }

                    Mock -CommandName Get-SPUsageService -MockWith {
                        return @{
                            UsageLogCutTime     = $testParams.UsageLogCutTime
                            UsageLogDir         = $testParams.UsageLogLocation
                            UsageLogMaxFileSize = ($testParams.UsageLogMaxFileSizeKB * 1024)
                            UsageLogMaxSpaceGB  = 1
                        }
                    }
                }

                It "Should return false when the Test method is called" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should call the update service app cmdlet from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Set-SPUsageService
                }
            }

            Context -Name "When the service app exists but it shouldn't" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name   = "Test App"
                        Ensure = "Absent"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            TypeName      = "Usage and Health Data Collection Service Application"
                            DisplayName   = $testParams.Name
                            Name          = $testParams.Name
                            UsageDatabase = @{
                                Name                 = "db"
                                NormalizedDataSource = "server"
                            }
                        }
                        $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                            return @{ FullName = $getTypeFullName }
                        } -PassThru -Force
                        return $spServiceApp
                    }
                }

                It "Should return present from the Get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Present"
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should remove the service application in the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled Remove-SPServiceApplication
                }
            }

            Context -Name "When the service app doesn't exist and shouldn't" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name   = "Test App"
                        Ensure = "Absent"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        return $null
                    }
                }

                It "Should return absent from the Get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Absent"
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "The proxy for the service app is offline when it should be running" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name   = "Test App"
                        Ensure = "Present"
                    }

                    Mock -CommandName Get-SPServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            TypeName      = "Usage and Health Data Collection Service Application"
                            DisplayName   = $testParams.Name
                            Name          = $testParams.Name
                            UsageDatabase = @{
                                Name                 = "db"
                                NormalizedDataSource = "server"
                            }
                        }
                        $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                            return @{ FullName = $getTypeFullName }
                        } -PassThru -Force
                        return $spServiceApp
                    }

                    Mock -CommandName Get-SPServiceApplicationProxy -MockWith {
                        $proxy = (New-Object -TypeName "Object" |
                                Add-Member -MemberType ScriptMethod `
                                    -Name Provision `
                                    -Value {
                                    $Global:SPDscUSageAppProxyStarted = $true
                                } -PassThru |
                                    Add-Member -NotePropertyName Status `
                                        -NotePropertyValue "Disabled" `
                                        -PassThru |
                                        Add-Member -NotePropertyName TypeName `
                                            -NotePropertyValue "Usage and Health Data Collection Proxy" `
                                            -PassThru)
                        $proxy = $proxy | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                            return @{ FullName = $getTypeFullNameProxy }
                        } -PassThru -Force
                        return $proxy
                    }
                }

                It "Should return absent from the get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Absent"
                }

                It "Should return false from the test method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should start the proxy in the set method" {
                    $Global:SPDscUSageAppProxyStarted = $false
                    Set-TargetResource @testParams
                    $Global:SPDscUSageAppProxyStarted | Should -Be $true
                }
            }

            Context -Name "Running ReverseDsc Export" -Fixture {
                BeforeAll {
                    Import-Module (Join-Path -Path (Split-Path -Path (Get-Module SharePointDsc -ListAvailable).Path -Parent) -ChildPath "Modules\SharePointDSC.Reverse\SharePointDSC.Reverse.psm1")

                    Mock -CommandName Write-Host -MockWith { }

                    Mock -CommandName Get-TargetResource -MockWith {
                        return @{
                            Name                  = "Usage Service Application"
                            DatabaseName          = "SP_Usage"
                            DatabaseServer        = "SQL01"
                            UsageLogCutTime       = 5
                            UsageLogLocation      = "L:\UsageLogs"
                            UsageLogMaxFileSizeKB = 1024
                            UsageLogMaxSpaceGB    = 5
                            Ensure                = "Present"
                        }
                    }

                    Mock -CommandName Get-SPUsageApplication -MockWith {
                        $spServiceApp = @(
                            @{
                                TypeName    = "Usage Service Application"
                                DisplayName = "Usage Service Application"
                                Name        = "Usage Service Application"
                            }
                        )
                        return $spServiceApp
                    }

                    if ($null -eq (Get-Variable -Name 'spFarmAccount' -ErrorAction SilentlyContinue))
                    {
                        $mockPassword = ConvertTo-SecureString -String "password" -AsPlainText -Force
                        $Global:spFarmAccount = New-Object -TypeName System.Management.Automation.PSCredential ("contoso\spfarm", $mockPassword)
                    }

                    $result = @'
        SPUsageApplication UsageServiceApplication
        {
            DatabaseName          = "SP_Usage";
            DatabaseServer        = $ConfigurationData.NonNodeData.DatabaseServer;
            Ensure                = "Present";
            Name                  = "Usage Service Application";
            PsDscRunAsCredential  = $Credsspfarm;
            UsageLogCutTime       = 5;
            UsageLogLocation      = $ConfigurationData.NonNodeData.UsageLogLocation;
            UsageLogMaxFileSizeKB = 1024;
            UsageLogMaxSpaceGB    = 5;
        }

'@
                }

                It "Should return valid DSC block from the Export method" {
                    Export-TargetResource | Should -Be $result
                }
            }
        }
    }
}
finally
{
    Invoke-TestCleanup
}
