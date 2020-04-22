$script:resourceModulePath = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$script:modulesFolderPath = Join-Path -Path $script:resourceModulePath -ChildPath 'Modules'
$script:resourceHelperModulePath = Join-Path -Path $script:modulesFolderPath -ChildPath 'SharePointDsc.Util'
Import-Module -Name (Join-Path -Path $script:resourceHelperModulePath -ChildPath 'SharePointDsc.Util.psm1')

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ServiceAppName,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Context,

        [Parameter()]
        [System.String[]]
        $QueryConditionNames,

        [Parameter()]
        [System.String[]]
        $QueryConditions,

        [Parameter()]
        [System.String]
        $QueryTemplate,

        [Parameter()]
        [System.Boolean]
        $AlwaysShow,

        [Parameter()]
        [System.Int]
        $RowLimit,

        [Parameter()]
        [System.String]
        $ResultTitleUrl,

        [Parameter()]
        [System.String]
        $GroupTemplateId,

        [Parameter()]
        [System.String]
        $BestBetsActionUrl,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Change ranked results by changing the query", "Promoted Results", "Result Blocks")]
        [System.String]
        $ResultType,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Getting query rule '$Name'"

    $result = Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments @($PSBoundParameters) `
        -ScriptBlock {
        $params = $args[0]

        $ssa = Get-SPEnterpriseSearchServiceApplication -Identity $params.ServiceAppName
        if ($null -eq $ssa)
        {
            throw ("The specified Search Service Application $($params.ServiceAppName) is " + `
                    "invalid. Please make sure you specify the name of an existing service application.")
        }

        $federationManager = New-Object Microsoft.Office.Server.Search.Administration.Query.FederationManager($ssa)
        $scope = Get-SPEnterpriseSearchOwner -Level SSA
        $queryRuleManager = New-Object Microsoft.Office.Server.Search.Query.Rules.QueryRuleManager($ssa)
        $searchObjectFilter = New-Object Microsoft.Office.Server.Search.Administration.SearchObjectFilter($scope)

        $queryRule = $queryRuleManager.GetQueryRules($searchObjectFilter) | Where-Object -FilterScript {
            $_.DisplayName -eq $params.DisplayName
        }

        if ($null -eq $queryRule)
        {
            return @{
                ServiceAppName = $params.ServiceAppName
                Name           = $params.Name
                DisplayName    = $params.DisplayName
                Context        = $params.Context
                Ensure         = "Absent"
            }
        }
        else
        {
            $results = @{
                ServiceAppName = $params.ServiceAppName
                Name           = $params.Name
                DisplayName    = $params.DisplayName
                Context        = $params.Context
                Ensure         = "Present"
            }
            return $results
        }
    }
    return $result
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ServiceAppName,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Context,

        [Parameter()]
        [System.String[]]
        $QueryConditionNames,

        [Parameter()]
        [System.String[]]
        $QueryConditions,

        [Parameter()]
        [System.String]
        $QueryTemplate,

        [Parameter()]
        [System.Boolean]
        $AlwaysShow,

        [Parameter()]
        [System.Int]
        $RowLimit,

        [Parameter()]
        [System.String]
        $ResultTitleUrl,

        [Parameter()]
        [System.String]
        $GroupTemplateId,

        [Parameter()]
        [System.String]
        $BestBetsActionUrl,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Change ranked results by changing the query", "Promoted Results", "Result Blocks")]
        [System.String]
        $ResultType,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Setting Managed Property Setting for '$Name'"

    # Obtain information about the current state of the Managed property (if it exists)
    $CurrentValues = Get-TargetResource @PSBoundParameters

    # Validate that the specified crawled properties are all valid and existing
    Invoke-SPDscCommand -Credential $InstallAccount `
        -Arguments @($PSBoundParameters, `
            $CurrentValues) `
        -ScriptBlock {
        $params = $args[0]
        $CurrentValues = $args[1]

        $ssa = Get-SPEnterpriseSearchServiceApplication -Identity $params.ServiceAppName
        if ($null -eq $ssa)
        {
            throw ("The specified Search Service Application $($params.ServiceAppName) is " + `
                    "invalid. Please make sure you specify the name of an existing service application.")
        }

        $needToRecreate = $false

        $federationManager = New-Object Microsoft.Office.Server.Search.Administration.Query.FederationManager($ssa)
        $scope = Get-SPEnterpriseSearchOwner -Level SSA
        $queryRuleManager = New-Object Microsoft.Office.Server.Search.Query.Rules.QueryRuleManager($ssa)
        $searchObjectFilter = New-Object Microsoft.Office.Server.Search.Administration.SearchObjectFilter($scope)

        $queryRule = $queryRuleManager.GetQueryRules($searchObjectFilter) | Where-Object -FilterScript {
            $_.DisplayName -eq $params.DisplayName
        }

        if ($params.Ensure -eq "Absent" -and `
                $CurrentValues.Ensure -eq "Present" -or `
            ($params.DisplayName -ne $CurrentValues.DisplayName -and `
                    $CurrentValues.Ensure -eq "Present"))
        {
            $queryRule = $queryRuleManager.GetQueryRules($searchObjectFilter) | Where-Object -FilterScript {
                $_.DisplayName -eq $params.DisplayName
            }

            if ($params.Context -ne $CurrentValues.Context)
            {
                Write-Verbose "Detected a change to context from $($CurrentValues.Context) to `
                               $($params.Context)"
                $needToRecreate = $true
            }
        }

        if (($CurrentValues.Ensure -eq "Absent" -and $params.Ensure -eq "Present") -or $needToRecreate)
        {
            $ruleName = $params.Name

            $newQueryRule = $queryRules.CreateQueryRule($ruleName, $null, $null, $true)

            $queryContext = $params.Context

            $resultSource = Get-SPEnterpriseSearchResultSource -SearchApplication $ssa -Owner $scope | Where-Object { $_.Name -eq $queryContext }

            if ($null -eq $resultSource)
            {
                throw ("The result source $($queryContext) is " + `
                        "invalid. Please make sure you specify the name of an existing result source.")
            }

            $querySourceContextCondition = $newQueryRule.CreateSourceContextCondition($resultSource)

            $newQueryConditions = $params.QueryConditions

            $i = 0

            $conditions = $newQueryRule.QueryConditions
            $newQueryConditionNames = $params.QueryConditionName

            foreach ($newQueryConditionName in $newQueryConditionNames)
            {
                switch ($newQueryConditionName.Trim())
                {
                    "Query Matches Keyword Exactly"
                    {
                        [string[]]$newQueryCondition = @($newQueryConditions[$i])
                        $Parameter = $conditions.CreateKeywordCondition($newQueryCondition, $true)
                        $Parameter.MatchingOptions = "FullQuery"
                        $Parameter.SubjectTermsOrigin = "MatchedTerms"
                        break
                    }
                    "Query Contains Action Term"
                    {
                        [string[]]$newQueryCondition = @($newQueryConditions[$i])
                        $Parameter = $conditions.CreateKeywordCondition($newQueryCondition, $true)
                        $Parameter.SubjectTermsOrigin = "Remainder"
                        $Parameter.MatchingOptions = "ProperPrefix, ProperSuffix"
                        break
                    }
                    "Query More Common in Source"
                    {
                        [string[]]$newQueryCondition = @($newQueryConditions[$i])
                        $Parameter = $conditions.CreateRegularExpressionCondition($newQueryCondition, $true)
                        break
                    }
                    "Result Type Commonly Clicked"
                    {
                        $Parameter = $conditions.CreateCommonQueryCondition($federationManager.GetDefaultSource($scope).Id, $true)
                        break
                    }
                    default
                    {
                        [string[]]$newQueryCondition = @($newQueryConditions[$i])
                        $Parameter = $conditions.CreateKeywordCondition($newQueryCondition, $true)
                        $Parameter.MatchingOptions = "FullQuery"
                        $Parameter.SubjectTermsOrigin = "MatchedTerms"
                    }
                }
                $i++
            }

            $queryTemplate = $params.QueryTemplate
            $resultTitle = $params.ResultTitle
            $alwaysShow = $params.AlwaysShow
            $rowLimit = $params.RowLimit
            $resultTitleUrl = $params.ResultTitleUrl
            $groupTemplateId = $params.GroupTemplateId
            $bestBetsActionUrl = $params.BestBetsActionUrl
            $actionType = $params.ResultType

            switch ($actionType.Trim())
            {
                "Change ranked results by changing the query"
                {
                    $QueryRuleAction = $newQueryRule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::ChangeQuery)
                    $QueryRuleAction.QueryTransform.OverrideProperties = New-Object Microsoft.Office.Server.Search.Query.Rules.QueryTransformProperties
                    $QueryRuleAction.QueryTransform.SourceId = $resultSource.Id
                    $newQueryRule.ChangeQueryAction.QueryTransform.QueryTemplate = $queryTemplate
                    break
                }
                "Promoted Results"
                {
                    $QueryRuleAction = $newQueryRule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::AssignBestBet)
                    $Uri = New-Object System.Uri($bestBetsActionUrl, $true)
                    $BestBetsCollection = $QueryRuleManager.GetBestBets($SearchObjectFilter)

                    if ($BestBetsCollection.Contains($Uri))
                    {
                        $bestBetsEnumeration = $BestBetsCollection.GetEnumerator()
                        $bestBetsEnumeration.Reset()

                        while ($bestBetsEnumeration.MoveNext())
                        {
                            $bestBets = $bestBetsEnumeration.Current
                            if ($bestBets.Url.OriginalString -eq $Uri.OriginalString)
                            {
                                $newQueryRule.AssignBestBetsAction.BestBetIds.Add($bestBets.Id)
                            }
                        }
                    }
                    else
                    {
                        throw ("Key word with url '$bestBetsActionUrl' does not exist.")
                    }
                    break
                }
                "Result Blocks"
                {
                    $QueryRuleAction = $newQueryRule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::CreateResultBlock)
                    $QueryRuleAction.QueryTransform.OverrideProperties = New-Object Microsoft.Office.Server.Search.Query.Rules.QueryTransformProperties
                    $QueryRuleAction.QueryTransform.SourceId = $resultSource.Id
                    $QueryRuleAction.QueryTransform.QueryTemplate = $queryTemplate
                    $QueryRuleAction.AlwaysShow = $alwaysShow
                    $QueryRuleAction.QueryTransform.OverrideProperties["RowLimit"] = $rowLimit
                    $QueryRuleAction.ResultTitleUrl = $resultTitleUrl
                    $QueryRuleAction.GroupTemplateId = $groupTemplateId
                    $QueryRuleAction.ResultTitle.Add(1031, $resultTitle) #TODO -language
                    break
                }

                default
                {
                    $QueryRuleAction = $newQueryRule.CreateQueryAction([Microsoft.Office.Server.Search.Query.Rules.QueryActionType]::ChangeQuery)
                    $QueryRuleAction.QueryTransform.OverrideProperties = New-Object Microsoft.Office.Server.Search.Query.Rules.QueryTransformProperties
                    $QueryRuleAction.QueryTransform.SourceId = $resultSource.Id
                    $newQueryRule.ChangeQueryAction.QueryTransform.QueryTemplate = $queryTemplate
                }
            }
            $newQueryRule.Update()
        }
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ServiceAppName,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Context,

        [Parameter()]
        [System.String[]]
        $QueryConditionNames,

        [Parameter()]
        [System.String[]]
        $QueryConditions,

        [Parameter()]
        [System.String]
        $QueryTemplate,

        [Parameter()]
        [System.Boolean]
        $AlwaysShow,

        [Parameter()]
        [System.Int]
        $RowLimit,

        [Parameter()]
        [System.String]
        $ResultTitleUrl,

        [Parameter()]
        [System.String]
        $GroupTemplateId,

        [Parameter()]
        [System.String]
        $BestBetsActionUrl,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Change ranked results by changing the query", "Promoted Results", "Result Blocks")]
        [System.String]
        $ResultType,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $InstallAccount
    )

    Write-Verbose -Message "Testing Query Rule Setting for '$Name'"

    $PSBoundParameters.Ensure = $Ensure

    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-SPDscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-SPDscHashtableToString -Hashtable $PSBoundParameters)"

    return Test-SPDscParameterState -CurrentValues $CurrentValues `
        -Source $($MyInvocation.MyCommand.Source) `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck @("Name",
        "DisplayName",
        "Context",
        "Ensure")
}

Export-ModuleMember -Function *-TargetResource
