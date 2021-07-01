[CmdletBinding()]
param
(
    [Parameter()]
    [string]
    $SharePointCmdletModule = (Join-Path -Path $PSScriptRoot `
            -ChildPath "..\Stubs\SharePoint\15.0.4805.1000\Microsoft.SharePoint.PowerShell.psm1" `
            -Resolve)
)

$script:DSCModuleName = 'SharePointDsc'
$script:DSCResourceName = 'SPProjectServerServiceApp'
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
                Invoke-Command -ScriptBlock $Global:SPDscHelper.InitializeScript -NoNewScope

                if ($Global:SPDscHelper.CurrentStubBuildNumber.Major -eq 16)
                {
                    # Initialize Tests
                    $getTypeFullName = "Microsoft.Office.Project.Server.Administration.PsiServiceApplication"

                    # Mocks for all contexts
                    Mock -CommandName New-SPProjectServiceApplication -MockWith { }
                    Mock -CommandName Set-SPProjectServiceApplication -MockWith { }
                    Mock -CommandName Remove-SPServiceApplication -MockWith { }
                    Mock -CommandName New-SPProjectServiceApplicationProxy -MockWith { }

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
            }

            switch ($Global:SPDscHelper.CurrentStubBuildNumber.Major)
            {
                15
                {
                    Context -Name "All methods throw exceptions as Project Server support in SharePointDsc is only for 2016" -Fixture {
                        It "Should throw on the get method" {
                            { Get-TargetResource @testParams } | Should -Throw
                        }

                        It "Should throw on the test method" {
                            { Test-TargetResource @testParams } | Should -Throw
                        }

                        It "Should throw on the set method" {
                            { Set-TargetResource @testParams } | Should -Throw
                        }
                    }
                }
                16
                {
                    # Test contexts
                    Context -Name "When no service applications exist in the current farm" -Fixture {
                        BeforeAll {
                            $testParams = @{
                                Name            = "Test Project Server App"
                                ApplicationPool = "Test App Pool"
                                Ensure          = "Present"
                            }

                            Mock -CommandName Get-SPServiceApplication -MockWith {
                                return $null
                            }
                        }

                        It "Should return absent from the Get method" {
                            (Get-TargetResource @testParams).Ensure | Should -Be "Absent"
                        }

                        It "Should return false when the Test method is called" {
                            Test-TargetResource @testParams | Should -Be $false
                        }

                        It "Should create a new service application in the set method" {
                            Set-TargetResource @testParams
                            Assert-MockCalled -CommandName New-SPProjectServiceApplication
                        }
                    }

                    Context -Name "When service applications exist in the current farm but the specific Project Server app does not" -Fixture {
                        BeforeAll {
                            $testParams = @{
                                Name            = "Test Project Server App"
                                ApplicationPool = "Test App Pool"
                                Ensure          = "Present"
                            }

                            Mock -CommandName Get-SPServiceApplication -MockWith {
                                $spServiceApp = [PSCustomObject]@{
                                    DisplayName = $testParams.Name
                                    Name        = $testParams.Name
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
                    }

                    Context -Name "When a service application exists and is configured correctly" -Fixture {
                        BeforeAll {
                            $testParams = @{
                                Name            = "Test Project Server App"
                                ApplicationPool = "Test App Pool"
                                Ensure          = "Present"
                            }

                            Mock -CommandName Get-SPServiceApplication -MockWith {
                                $spServiceApp = [PSCustomObject]@{
                                    TypeName        = "Project Application Services"
                                    DisplayName     = $testParams.Name
                                    Name            = $testParams.Name
                                    ApplicationPool = @{ Name = $testParams.ApplicationPool }
                                }
                                $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                                    return @{ FullName = $getTypeFullName }
                                } -PassThru -Force
                                return $spServiceApp
                            }
                        }

                        It "Should return present from the get method" {
                            (Get-TargetResource @testParams).Ensure | Should -Be "Present"
                        }

                        It "Should return true when the Test method is called" {
                            Test-TargetResource @testParams | Should -Be $true
                        }
                    }

                    Context -Name "When a service application exists and is not configured correctly" -Fixture {
                        BeforeAll {
                            $testParams = @{
                                Name            = "Test Project Server App"
                                ApplicationPool = "Test App Pool"
                                Ensure          = "Present"
                            }

                            Mock -CommandName Get-SPServiceApplication -MockWith {
                                $spServiceApp = [PSCustomObject]@{
                                    TypeName        = "Project Application Services"
                                    DisplayName     = $testParams.Name
                                    Name            = $testParams.Name
                                    ApplicationPool = @{ Name = "Wrong App Pool Name" }
                                }
                                $spServiceApp = $spServiceApp | Add-Member -MemberType ScriptMethod -Name GetType -Value {
                                    return @{ FullName = $getTypeFullName }
                                } -PassThru -Force
                                return $spServiceApp
                            }
                            Mock -CommandName Get-SPServiceApplicationPool {
                                return @{
                                    Name = $testParams.ApplicationPool
                                }
                            }
                        }

                        It "Should return false when the Test method is called" {
                            Test-TargetResource @testParams | Should -Be $false
                        }

                        It "Should call the update service app cmdlet from the set method" {
                            Set-TargetResource @testParams

                            Assert-MockCalled Set-SPProjectServiceApplication
                        }
                    }

                    Context -Name "When the service app exists but it shouldn't" -Fixture {
                        BeforeAll {
                            $testParams = @{
                                Name            = "Test App"
                                ApplicationPool = "-"
                                Ensure          = "Absent"
                            }

                            Mock -CommandName Get-SPServiceApplication -MockWith {
                                $spServiceApp = [PSCustomObject]@{
                                    TypeName        = "Project Application Services"
                                    DisplayName     = $testParams.Name
                                    Name            = $testParams.Name
                                    ApplicationPool = @{ Name = $testParams.ApplicationPool }
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
                                Name            = "Test App"
                                ApplicationPool = "-"
                                Ensure          = "Absent"
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
                }
            }
        }
    }
}
finally
{
    Invoke-TestCleanup
}
