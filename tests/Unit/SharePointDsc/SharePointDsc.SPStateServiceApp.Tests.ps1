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
$script:DSCResourceName = 'SPStateServiceApp'
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
                $mockPassword = ConvertTo-SecureString -String "password" -AsPlainText -Force
                $mockCredential = New-Object -TypeName System.Management.Automation.PSCredential `
                    -ArgumentList @("username", $mockPassword)

                # Mocks for all contexts
                Mock -CommandName New-SPStateServiceDatabase -MockWith { return @{ } }
                Mock -CommandName New-SPStateServiceApplication -MockWith { return @{ } }
                Mock -CommandName New-SPStateServiceApplicationProxy -MockWith { return @{ } }
                Mock -CommandName Remove-SPServiceApplication -MockWith { }
            }

            # Test contexts
            Context -Name "the service app doesn't exist and should" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name                = "State Service App"
                        ProxyName           = "State Service Proxy"
                        DatabaseName        = "SP_StateService"
                        DatabaseServer      = "SQL.test.domain"
                        DatabaseCredentials = $mockCredential
                        Ensure              = "Present"
                    }

                    Mock -CommandName New-SPStateServiceApplication -MockWith {
                        $returnVal = @{
                            Name = $testParams.Name
                        }
                        $returnVal = $returnVal | Add-Member -MemberType ScriptMethod `
                            -Name IsConnected -Value {
                            return $true
                        } -PassThru

                        return $returnVal
                    }
                    Mock -CommandName Get-SPServiceApplication -MockWith { return $null }
                    Mock -CommandName Get-SPServiceApplicationProxy -MockWith {
                        $proxiesToReturn = @()
                        $proxy = @{
                            Name        = $testParams.ProxyName
                            DisplayName = $testParams.ProxyName
                        }
                        $proxy = $proxy | Add-Member -MemberType ScriptMethod `
                            -Name Delete `
                            -Value { } `
                            -PassThru
                        $proxiesToReturn += $proxy

                        return $proxiesToReturn
                    }
                }

                It "Should return absent from the get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Absent"
                }

                It "Should return false from the get method" {
                    Test-TargetResource @testParams | Should -Be $false
                }

                It "Should create a state service app from the set method" {
                    Set-TargetResource @testParams
                    Assert-MockCalled New-SPStateServiceApplication
                }
            }

            Context -Name "the service app exists and should" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name                = "State Service App"
                        DatabaseName        = "SP_StateService"
                        DatabaseServer      = "SQL.test.domain"
                        DatabaseCredentials = $mockCredential
                        Ensure              = "Present"
                    }

                    Mock -CommandName Get-SPStateServiceApplication -MockWith {
                        $returnVal = @{
                            DisplayName = $testParams.Name
                            Name        = $testParams.Name
                        }
                        $returnVal = $returnVal | Add-Member -MemberType ScriptMethod `
                            -Name IsConnected -Value {
                            return $true
                        } -PassThru

                        return $returnVal
                    }
                    Mock -CommandName Get-SPServiceApplicationProxy -MockWith {
                        $proxiesToReturn = @()
                        $proxy = @{
                            Name        = $testParams.ProxyName
                            DisplayName = $testParams.ProxyName
                        }
                        $proxy = $proxy | Add-Member -MemberType ScriptMethod `
                            -Name Delete `
                            -Value { } `
                            -PassThru
                        $proxiesToReturn += $proxy

                        return $proxiesToReturn
                    }
                }

                It "Should return present from the get method" {
                    (Get-TargetResource @testParams).Ensure | Should -Be "Present"
                }

                It "Should return true from the test method" {
                    Test-TargetResource @testParams | Should -Be $true
                }
            }

            Context -Name "When the service app exists but it shouldn't" -Fixture {
                BeforeAll {
                    $testParams = @{
                        Name         = "State Service App"
                        DatabaseName = "-"
                        Ensure       = "Absent"
                    }

                    Mock -CommandName Get-SPStateServiceApplication -MockWith {
                        return @{
                            DisplayName = $testParams.Name
                        }
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
                        Name         = "State Service App"
                        DatabaseName = "-"
                        Ensure       = "Absent"
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

            Context -Name "Running ReverseDsc Export" -Fixture {
                BeforeAll {
                    Import-Module (Join-Path -Path (Split-Path -Path (Get-Module SharePointDsc -ListAvailable).Path -Parent) -ChildPath "Modules\SharePointDSC.Reverse\SharePointDSC.Reverse.psm1")

                    Mock -CommandName Write-Host -MockWith { }

                    Mock -CommandName Get-TargetResource -MockWith {
                        return @{
                            Name           = "State Service Application"
                            ProxyName      = "State Service Application Proxy"
                            DatabaseName   = "SP_State"
                            DatabaseServer = "SQL01"
                            Ensure         = "Present"
                        }
                    }

                    Mock -CommandName Get-SPStateServiceApplication -MockWith {
                        $spServiceApp = [PSCustomObject]@{
                            DisplayName = "State Service Application"
                            Name        = "State Service Application"
                        }
                        return $spServiceApp
                    }

                    if ($null -eq (Get-Variable -Name 'spFarmAccount' -ErrorAction SilentlyContinue))
                    {
                        $mockPassword = ConvertTo-SecureString -String "password" -AsPlainText -Force
                        $Global:spFarmAccount = New-Object -TypeName System.Management.Automation.PSCredential ("contoso\spfarm", $mockPassword)
                    }

                    $result = @'
        SPStateServiceApp StateServiceApplication
        {
            DatabaseName         = "SP_State";
            DatabaseServer       = $ConfigurationData.NonNodeData.DatabaseServer;
            Ensure               = "Present";
            Name                 = "State Service Application";
            ProxyName            = "State Service Application Proxy";
            PsDscRunAsCredential = $Credsspfarm;
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
