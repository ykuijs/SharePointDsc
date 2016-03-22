function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [parameter(Mandatory = $true)]  [System.String] $WebAppUrl,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $Members,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $MembersToInclude,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $MembersToExclude,
        [parameter(Mandatory = $false)] [System.Boolean] $SetCacheAccountsPolicy,
        [parameter(Mandatory = $false)] [System.Management.Automation.PSCredential] $InstallAccount
    )

    if ($Members -and (($MembersToInclude) -or ($MembersToExclude))) {
        Write-Verbose -Verbose "Cannot use the Members parameter together with the MembersToInclude or MembersToExclude parameters"
        return $null
    }

    if (!$Members -and !$MembersToInclude -and !$MembersToExclude) {
        Write-Verbose -Verbose "At least one of the following parameters must be specified: Members, MembersToInclude, MembersToExclude"
        return $null
    }

    foreach ($member in $Members) {
        if (($member.ActAsSystemAccount -eq $true) -and ($member.PermissionLevel -ne "Full Control")) {
            Write-Verbose -Verbose "Members Parameter: You cannot specify ActAsSystemAccount with any other permission than Full Control"        
            return $null
        }
    }

    foreach ($member in $MembersToInclude) {
        if (($member.ActAsSystemAccount -eq $true) -and ($member.PermissionLevel -ne "Full Control")) {
            Write-Verbose -Verbose "MembersToInclude Parameter: You cannot specify ActAsSystemAccount with any other permission than Full Control"        
            return $null
        }
    }
    
    Write-Verbose -Message "Getting web app policy for $UserName at $WebAppUrl"

    $result = Invoke-xSharePointCommand -Credential $InstallAccount -Arguments $PSBoundParameters -ScriptBlock {
        $params = $args[0]

        $wa = Get-SPWebApplication -Identity $params.WebAppUrl -ErrorAction SilentlyContinue

        if ($null -eq $wa) { return $null }

        $members = @()
        foreach ($policy in $wa.Policies) {
            $member = @{}
            $member.Username = $policy.UserName
            $member.PermissionLevel = $policy.PolicyRoleBindings.Name
            $member.ActAsSystemAccount = $policy.IsSystemUser
            $members += $member
        }

        $returnval = @{
                WebAppUrl = $params.WebAppUrl
                Members = $members
                MembersToInclude = $params.MembersToInclude
                MembersToExclude = $params.MembersToExclude
                InstallAccount = $params.InstallAccount
        }

        if ($wa.Properties.ContainsKey("portalsuperuseraccount")) {
            $returnval.SuperUserAccount = $wa.Properties["portalsuperuseraccount"]
        }
        if ($wa.Properties.ContainsKey("portalsuperreaderaccount")) {
            $returnval.SuperReaderAccount = $wa.Properties["portalsuperreaderaccount"]
        }
        
        return $returnval
    }
    return $result
}


function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory = $true)]  [System.String] $WebAppUrl,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $Members,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $MembersToInclude,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $MembersToExclude,
        [parameter(Mandatory = $false)] [System.Boolean] $SetCacheAccountsPolicy,
        [parameter(Mandatory = $false)] [System.Management.Automation.PSCredential] $InstallAccount
    )

    Write-Verbose -Message "Setting web app policy for $UserName at $WebAppUrl"

    if ($Members -and (($MembersToInclude) -or ($MembersToExclude))) {
        Throw "Cannot use the Members parameter together with the MembersToInclude or MembersToExclude parameters"
    }

    if (!$Members -and !$MembersToInclude -and !$MembersToExclude) {
        throw "At least one of the following parameters must be specified: Members, MembersToInclude, MembersToExclude"
    }

    foreach ($member in $Members) {
        if (($member.ActAsSystemAccount -eq $true) -and ($member.PermissionLevel -ne "Full Control")) {
            throw "Members Parameter: You cannot specify ActAsSystemAccount with any other permission than Full Control"        
        }
    }

    foreach ($member in $MembersToInclude) {
        if (($member.ActAsSystemAccount -eq $true) -and ($member.PermissionLevel -ne "Full Control")) {
            throw "MembersToInclude Parameter: You cannot specify ActAsSystemAccount with any other permission than Full Control"        
        }
    }
    
    $result = Invoke-xSharePointCommand -Credential $InstallAccount -Arguments @($PSBoundParameters,$PSScriptRoot) -ScriptBlock {
		$params = $args[0]
        $ScriptRoot = $args[1]

        $wa = Get-SPWebApplication -Identity $params.WebAppUrl -ErrorAction SilentlyContinue

        if ($null -eq $wa) { return $null }

        $denyAll     = $wa.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::DenyAll)
		$denyWrite   = $wa.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::DenyWrite)
        $fullControl = $wa.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl)
        $fullRead    = $wa.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead)

        if (($SetCacheAccountsPolicy) -and (-not($wa.Properties.ContainsKey("portalsuperuseraccount")) -or -not($wa.Properties.ContainsKey("portalsuperreaderaccount")))) {
            throw "Cache accounts not configured properly. PortalSuperUserAccount or PortalSuperReaderAccount properties is not configured."
        }
        
		$members = @()
        foreach ($policy in $wa.Policies) {
            $member = @{}
            $member.Username = $policy.UserName
            $member.PermissionLevel = $policy.PolicyRoleBindings.Name
            $member.ActAsSystemAccount = $policy.IsSystemUser
            $members += $member
        }

		if ($params.Members) {
            Write-Verbose -Verbose "Processing Members parameter"

            $allMembers = @()
            foreach ($member in $params.Members) {
                $allMembers += $member
            }

            if ($params.SetCacheAccountsPolicy) {
                Write-Verbose -Verbose "Adding Cache Accounts to Members parameter"
                $psuAccount = @{
                    UserName = $wa.Properties["portalsuperuseraccount"]
                    PermissionLevel = "Full Read"
                }
                $allMembers += $psuAccount
                
                $psrAccount = @{
                    UserName = $wa.Properties["portalsuperreaderaccount"]
                    PermissionLevel = "Full Read"
                }
                $allMembers += $psrAccount
            }

            Import-Module (Join-Path $ScriptRoot "..\..\Modules\xSharePoint.WebAppPolicy\xSPWebAppPolicy.psm1" -Resolve)
			$differences = ComparePolicies $members $allMembers

			foreach ($difference in $differences) {
				$user = $difference.Keys[0]
				$change = $difference[$user]
				switch ($change) {
					Additional
						{
							## Policy contains additional account, remove this account
							Write-Verbose -Verbose "Removing $user"
                            Remove-WebAppPolicy $wa.Policies $user
						}
					Different
						{
							## Account exists but has the incorrect settings, correct this account
							Write-Verbose -Verbose "Changing $user"
							$policy = $wa.Policies | Where-Object { $_.UserName -eq $user }
							$usersettings = GetUserFromCollection $allMembers $user
							if ($usersettings.ActAsSystemAccount -ne $policy.IsSystemUser) { $policy.IsSystemUser = $usersettings.ActAsSystemAccount }
							
                            $polbinddiff = Compare-Object -ReferenceObject $policy.PolicyRoleBindings.Name -DifferenceObject $usersettings.PermissionLevel
                            if ($polbinddiff -ne $null) {
                                $policy.PolicyRoleBindings.RemoveAll()
                                foreach ($permissionLevel in $usersettings.PermissionLevel) {
                                    switch ($permissionLevel) {
                                        "Deny All" {
                                            $policy.PolicyRoleBindings.Add($denyAll)
                                        }
                                        "Deny Write" {
                                            $policy.PolicyRoleBindings.Add($denyWrite)
                                        }
                                        "Full Control" {
                                            $policy.PolicyRoleBindings.Add($fullControl)
                                        }
                                        "Full Read" {
                                            $policy.PolicyRoleBindings.Add($fullRead)
                                        }
                                    }
                                }
							}
						}
					Missing
						{
							## Account is missing, add this account
							Write-Verbose -Verbose "Adding $user"
							$usersettings = GetUserFromCollection $allMembers $user
							$newPolicy = $wa.Policies.Add($user, $user)
                            foreach ($permissionLevel in $usersettings.PermissionLevel) {
                                switch ($permissionLevel) {
                                    "Deny All" {
                                        $newPolicy.PolicyRoleBindings.Add($denyAll)
                                    }
                                    "Deny Write" {
                                        $newPolicy.PolicyRoleBindings.Add($denyWrite)
                                    }
                                    "Full Control" {
                                        $newPolicy.PolicyRoleBindings.Add($fullControl)
                                    }
                                    "Full Read" {
                                        $newPolicy.PolicyRoleBindings.Add($fullRead)
                                    }
                                }
                            }
							if ($usersettings.ActAsSystemAccount) {
								$newPolicy.IsSystemUser = $usersettings.ActAsSystemAccount
							}
						}
				}
                $wa.Update()
			}
		}

		if ($params.MembersToInclude) {
            Write-Verbose -Verbose "Processing MembersToInclude parameter"

            $allMembers = @()
            foreach ($member in $MembersToInclude) {
                $allMembers += $member
            }

            if ($params.SetCacheAccountsPolicy) {
                Write-Verbose -Verbose "Adding Cache Accounts to MembersToInclude parameter"
                $psuAccount = @{
                    UserName = $wa.Properties["portalsuperuseraccount"]
                    PermissionLevel = "Full Read"
                }
                $allMembers += $psuAccount
                
                $psrAccount = @{
                    UserName = $wa.Properties["portalsuperreaderaccount"]
                    PermissionLevel = "Full Read"
                }
                $allMembers += $psrAccount
            }
            
            Import-Module (Join-Path $ScriptRoot "..\..\Modules\xSharePoint.WebAppPolicy\xSPWebAppPolicy.psm1" -Resolve)
            
			foreach ($member in $allMembers) {
                $policy = $wa.Policies | Where-Object { $_.UserName -eq $member.UserName }

                if ($policy -ne $null) {
                    # User exists. Check permissions
                    Write-Verbose -Verbose "User $($member.UserName) exists, checking permissions"
                    if ($member.ActAsSystemAccount -ne $policy.IsSystemUser) { $policy.IsSystemUser = $member.ActAsSystemAccount }

                    $polbinddiff = Compare-Object -ReferenceObject $policy.PolicyRoleBindings.Name -DifferenceObject $member.PermissionLevel
                    if ($polbinddiff -ne $null) {
                        $policy.PolicyRoleBindings.RemoveAll()
                        foreach ($permissionLevel in $member.PermissionLevel) {
                            switch ($permissionLevel) {
                                "Deny All" {
                                    $policy.PolicyRoleBindings.Add($denyAll)
                                }
                                "Deny Write" {
                                    $policy.PolicyRoleBindings.Add($denyWrite)
                                }
                                "Full Control" {
                                    $policy.PolicyRoleBindings.Add($fullControl)
                                }
                                "Full Read" {
                                    $policy.PolicyRoleBindings.Add($fullRead)
                                }
                            }
                        }
                    }
                } else {
                    # User does not exist. Add user
                    Write-Verbose -Verbose "Adding $($member.UserName)"
                    $newPolicy = $wa.Policies.Add($member.UserName, $member.UserName)
                    foreach ($permissionLevel in $member.PermissionLevel) {
                        switch ($permissionLevel) {
                            "Deny All" {
                                $newPolicy.PolicyRoleBindings.Add($denyAll)
                            }
                            "Deny Write" {
                                $newPolicy.PolicyRoleBindings.Add($denyWrite)
                            }
                            "Full Control" {
                                $newPolicy.PolicyRoleBindings.Add($fullControl)
                            }
                            "Full Read" {
                                $newPolicy.PolicyRoleBindings.Add($fullRead)
                            }
                        }
                    }
                    if ($member.ActAsSystemAccount) {
                        $newPolicy.IsSystemUser = $member.ActAsSystemAccount
                    }
                }
                $wa.Update()
			}
		}

		if ($params.MembersToExclude) {
            Write-Verbose -Verbose "Processing MembersToExclude parameter"

            $psuAccount = $wa.Properties["portalsuperuseraccount"]
            $psrAccount = $wa.Properties["portalsuperreaderaccount"]

            Import-Module (Join-Path $ScriptRoot "..\..\Modules\xSharePoint.WebAppPolicy\xSPWebAppPolicy.psm1" -Resolve)
            
			foreach ($member in $params.MembersToExclude) {
                if (($psuAccount -eq $member.UserName) -or ($psrAccount -eq $member.UserName)) {
                    throw "You cannot exclude the Cache accounts from the Web Application Policy"
                }

                $policy = $wa.Policies | Where-Object { $_.UserName -eq $member.UserName }

                if ($policy -ne $null) {
                    # User exists. Delete user
                    Write-Verbose -Verbose "User $($member.UserName) exists, deleting"
                    Remove-WebAppPolicy $wa.Policies $member.UserName
                }
                $wa.Update()
			}
		}
    }
}


function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [parameter(Mandatory = $true)]  [System.String] $WebAppUrl,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $Members,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $MembersToInclude,
        [parameter(Mandatory = $false)] [Microsoft.Management.Infrastructure.CimInstance[]] $MembersToExclude,
        [parameter(Mandatory = $false)] [System.Boolean] $SetCacheAccountsPolicy,
        [parameter(Mandatory = $false)] [System.Management.Automation.PSCredential] $InstallAccount
    )

    $CurrentValues = Get-TargetResource @PSBoundParameters
    Write-Verbose -Message "Testing web app policy for $UserName at $WebAppUrl"
    if ($null -eq $CurrentValues) { return $false }

    if (($SetCacheAccountsPolicy) -and (-not($CurrentValues.ContainsKey("SuperUserAccount")) -or -not($CurrentValues.ContainsKey("SuperReaderAccount")))) {
        throw "Cache accounts not configured properly. PortalSuperUserAccount or PortalSuperReaderAccount properties is not configured."
    }
    
    if ($Members) {
        Write-Verbose "Processing Members - Start Test"
        
        $allMembers = @()
        foreach ($member in $Members) {
            $allMembers += $member
        }

        if ($SetCacheAccountsPolicy) {
            Write-Verbose "SetCacheAccountsPolicy is True. Adding Cache Accounts to list"
            $psuAccount = @{
                UserName = $CurrentValues.SuperUserAccount
                PermissionLevel = "Full Read"
            }
            $allMembers += $psuAccount
            
            $psrAccount = @{
                UserName = $CurrentValues.SuperReaderAccount
                PermissionLevel = "Full Read"
            }
            $allMembers += $psrAccount
        }

        Import-Module (Join-Path $PsScriptRoot "..\..\Modules\xSharePoint.WebAppPolicy\xSPWebAppPolicy.psm1" -Resolve)
        $differences = ComparePolicies $CurrentValues.Members $allMembers

        if ($differences.Count -eq 0) { return $true } else { return $false }
    }

    if ($MembersToInclude) {
        Write-Verbose "Processing MembersToInclude - Start Test"

        $allMembers = @()
        foreach ($member in $MembersToInclude) {
            $allMembers += $member
        }

        if ($SetCacheAccountsPolicy) {
            Write-Verbose "SetCacheAccountsPolicy is True. Adding Cache Accounts to list"
            $psuAccount = @{
                UserName = $CurrentValues.SuperUserAccount
                PermissionLevel = "Full Read"
            }
            $allMembers += $psuAccount
            
            $psrAccount = @{
                UserName = $CurrentValues.SuperReaderAccount
                PermissionLevel = "Full Read"
            }
            $allMembers += $psrAccount
        }
        
        foreach ($member in $allMembers) {			
    		$match = $false
            foreach ($policy in $CurrentValues.Members) {
				if ($policy.Username.ToLower() -eq $member.Username.ToLower()) {
					$match = $true
                    if ($member.ActAsSystemAccount) {
	    				if ($policy.ActAsSystemAccount -ne $member.ActAsSystemAccount) { $match = $false }
                    }

                    $polbinddiff = Compare-Object -ReferenceObject $policy.PermissionLevel.ToLower() -DifferenceObject $member.PermissionLevel.ToLower()
                    if ($polbinddiff -ne $null) { $match = $false }
				}
			}
            if ($match -eq $false) { return $match }
		}
        return $true
    }

    if ($MembersToExclude) {
        Write-Verbose "Processing MembersToExclude - Start Test"
		foreach ($member in $MembersToExclude) {
            Write-Verbose "$($CurrentValues.SuperUserAccount) - $($CurrentValues.SuperReaderAccount) - $($member.Username)"
            
            if (($CurrentValues.SuperUserAccount -eq $member.Username) -or ($CurrentValues.SuperReaderAccount -eq $member.Username)) {
                throw "You cannot exclude the Cache accounts from the Web Application Policy"
            }
            
            foreach ($policy in $CurrentValues.Members) {
				if ($policy.Username.ToLower() -eq $member.Username.ToLower()) {
					return $false
				}
			}
		}
    	return $true
    }
}

Export-ModuleMember -Function *-TargetResource
