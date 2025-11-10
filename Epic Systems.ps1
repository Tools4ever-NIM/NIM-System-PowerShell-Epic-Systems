# version: 1.0
#
# Epic Systems.ps1 - Epic Systems
#
$Log_MaskableKeys = @(
    'password',
    "proxy_password"
)

$Global:User = [System.Collections.ArrayList]@()
$Global:UserInfo = [System.Collections.ArrayList]@()
$Global:UserInfo_UserIDs = [System.Collections.ArrayList]@()
$Global:UserInfo_UserRoleIDs = [System.Collections.ArrayList]@()
$Global:UserInfo_LinkedTemplatesConfig = [System.Collections.ArrayList]@()
$Global:UserInfo_UserSubtemplateIDs = [System.Collections.ArrayList]@()
$Global:UserInfo_LinkedProviderID = [System.Collections.ArrayList]@()

$Properties = @{
    User = @(
        @{ name = 'ID';              type = 'string';   objectfields = $null;             options = @('default','key') },
        @{ name = 'Name';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ExternalID';      type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'EnRolRank';       type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'Status';          type = 'string';   objectfields = $null;             options = @('default') }
    )
    UserGroup = @(
        @{ name = 'UserID';              type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'Group';              type = 'string';   objectfields = $null;             options = @('default') }
    )
    UserInfo = @(
        @{ name = 'UserID';            type = 'string';   objectfields = $null;             options = @('default','key') },
        @{ name = 'Name';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ContactComment';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'LDAPOverrideID';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'IsPasswordChangeRequired';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'IsActive';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'StartDate';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'EndDate';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'UserAlias';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'UserPhotoPath';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'Sex';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ReportGrouper1';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ReportGrouper2';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ReportGrouper3';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'CategoryReportGrouper1';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'CategoryReportGrouper2';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'CategoryReportGrouper3';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'CategoryReportGrouper4';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'CategoryReportGrouper5';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'CategoryReportGrouper6';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'Notes';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'InBasketClassifications';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'UserDirectoryPath';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ProviderAtLoginOption';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'UserComplexName';            type = 'object';   objectfields = @('FirstName','GivenNameInitials','MiddleName','LastName','LastNamePrefix','SpouseLastName','SpousePrefix','Suffix','AcademicTitle','PrimaryTitle','SpouseLastNameFirst','FatherName','GrandfatherName');             options = @('default') },
        @{ name = 'AuthenticationConfigurationID';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'BlockStatus';            type = 'object';   objectfields = @('IsBlocked','Reason','Comment');             options = @('default') },
        @{ name = 'EmployeeDemographics';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'PrimaryManager';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'UsersManagers';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'DefaultLoginDepartmentID';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'CustomUserDictionaries';            type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ExternalIdentifiers';            type = 'string';   objectfields = $null;             options = @('default') }

        #Object arrays
        #@{ name = 'UserIDs';            type = 'string';   objectfields = $null;             options = @('default') },
        #@{ name = 'LinkedTemplatesConfig';            type = 'string';   objectfields = $null;             options = @('default') },
        #@{ name = 'AppliedTemplateID';            type = 'string';   objectfields = $null;             options = @('default') },
        #@{ name = 'AvailableLinkableTemplates';            type = 'string';   objectfields = $null;             options = @('default') },
        #@{ name = 'UserSubtemplateIDs';            type = 'string';   objectfields = $null;             options = @('default') },
        #@{ name = 'UserRoleIDs';            type = 'string';   objectfields = $null;             options = @('default') },
        #@{ name = 'LinkedProviderID';            type = 'string';   objectfields = $null;             options = @('default') },
    )
    UserInfo_UserID = @(
        @{ name = 'UserID';              type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ID';              type = 'string';   objectfields = $null;             options = @('default') }
        @{ name = 'Type';              type = 'string';   objectfields = $null;             options = @('default') }
    )
    
    UserInfo_UserRoleID = @(
        @{ name = 'UserID';              type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ID';              type = 'string';   objectfields = $null;             options = @('default') }
        @{ name = 'Type';              type = 'string';   objectfields = $null;             options = @('default') }
    )
    UserInfo_LinkedTemplatesConfig = @(
        @{ name = 'UserID';              type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ID';              type = 'string';   objectfields = $null;             options = @('default') }
        @{ name = 'Type';              type = 'string';   objectfields = $null;             options = @('default') }
        @{ name = 'Area';              type = 'string';   objectfields = $null;             options = @('default') }
    )
    UserInfo_UserSubtemplateID = @(
        @{ name = 'UserID';              type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ID';              type = 'string';   objectfields = $null;             options = @('default') }
        @{ name = 'Type';              type = 'string';   objectfields = $null;             options = @('default') }
    )
    UserInfo_LinkedProviderID = @(
        @{ name = 'UserID';              type = 'string';   objectfields = $null;             options = @('default') },
        @{ name = 'ID';              type = 'string';   objectfields = $null;             options = @('default') }
        @{ name = 'Type';              type = 'string';   objectfields = $null;             options = @('default') }
    )
}

#
# System functions
#
function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"

    if ($Connection) {
        @(
            @{
                name = 'hostname'
                type = 'textbox'
                label = 'Hostname'
                description = ''
                value = 'customer.epic.com'
            }
            @{
                name = 'clientId'
                type = 'textbox'
                label = 'Client ID'
                label_indent = $true
                tooltip = 'EPIC Client ID'
                value = ''
            }
            @{
                name = 'username'
                type = 'textbox'
                label = 'Username'
                label_indent = $true
                tooltip = 'Username'
                value = ''
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                tooltip = 'Password'
                value = ''
            }
            @{
                name = 'use_proxy'
                type = 'checkbox'
                label = 'Use Proxy'
                description = 'Use Proxy server for requests'
                value = $false # Default value of checkbox item
            }
            @{
                name = 'proxy_address'
                type = 'textbox'
                label = 'Proxy Address'
                description = 'Address of the proxy server'
                value = 'http://127.0.0.1:8888'
                disabled = '!use_proxy'
                hidden = '!use_proxy'
            }
            @{
                name = 'use_proxy_credentials'
                type = 'checkbox'
                label = 'Use Proxy Credentials'
                description = 'Use credentials for proxy'
                value = $false
                disabled = '!use_proxy'
                hidden = '!use_proxy'
            }
            @{
                name = 'proxy_username'
                type = 'textbox'
                label = 'Proxy Username'
                label_indent = $true
                description = 'Username account'
                value = ''
                disabled = '!use_proxy_credentials'
                hidden = '!use_proxy_credentials'
            }
            @{
                name = 'proxy_password'
                type = 'textbox'
                password = $true
                label = 'Proxy Password'
                label_indent = $true
                description = 'User account password'
                value = ''
                disabled = '!use_proxy_credentials'
                hidden = '!use_proxy_credentials'
            }
            @{
                name = 'nr_of_retries'
                type = 'textbox'
                label = 'Max. number of retry attempts'
                description = ''
                value = 5
            }
            @{
                name = 'retryDelay'
                type = 'textbox'
                label = 'Seconds to wait for retry'
                description = ''
                value = 2
            }
            @{
                name = 'nr_of_threads'
                type = 'textbox'
                label = 'Max. number of simultaneous requests'
                description = ''
                value = 20
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                description = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                description = ''
                value = 1
            }
        )
    }

    if ($TestConnection) {
        
    }

    if ($Configuration) {
        @()
    }

    Log info "Done"
}

function Idm-OnUnload {
}

#
# Object CRUD functions
#

function Idm-UsersRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'User'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return 
        }

        # Refresh cache if needed
        if ($script:User.Count -eq 0) {
            $Global:User.Clear()
            $properties = Get-ClassProperties -Class $Class -IncludeHidden $true

            # Initial empty SearchStateContext
            $searchStateContext = @{
                Identifier   = ''
                ResumeInfo   = ''
                CriteriaHash = ''
            }
            $i = 0;
            do {
                # Build SearchStateContext XML
                $searchStateXml = @"
<SearchStateContext>
    <Identifier>$($searchStateContext.Identifier)</Identifier>
    <ResumeInfo>$($searchStateContext.ResumeInfo)</ResumeInfo>
    <CriteriaHash>$($searchStateContext.CriteriaHash)</CriteriaHash>
</SearchStateContext>
"@

                $splat = @{ 
                        SystemParams = $system_params
                        Class = $Class
                        Uri = '/interconnect-amcurprd-personnel-username/httplistener.ashx'
                        Action = 'urn:epicsystems-com:Core.2008-04.Services.GetUserRecords'
                        ActionResponse = 'urn:epicsystems-com:Core.2008-04.Services.GetUserRecords'
                        Body = ''
                        LogMessage = if($i -gt 0) { "Result Count: [$($i)]" }
                }


                # Update SOAP body
                $splat.Body = @"
<GetUserRecords xmlns="urn:epicsystems-com:Core.2008-04.Services">
    <ConnectionFunctionalType/>
    <SearchCriteria>
        <SearchString/>
    </SearchCriteria>
    $($searchStateXml)
    <UserID>$($system_params.username)</UserID>
</GetUserRecords>
"@

                # Send request
                $response = Execute-SOAPRequest @splat

                # Namespace manager
                $nsMgr = New-Object System.Xml.XmlNamespaceManager($response.NameTable)
                $nsMgr.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/")
                $nsMgr.AddNamespace("core", "urn:epicsystems-com:Core.2008-04.Services")

                # Extract and output records
                $records = $response.SelectNodes("//core:ResultRecord", $nsMgr)

                foreach ($rowItem in $records) {
                    $row = New-Object -TypeName PSObject -Property ([ordered]@{ })

                    foreach ($prop in $properties.GetEnumerator()) {
                        $row | Add-Member -NotePropertyName $prop -NotePropertyValue ""
                    }

                    foreach ($prop in $rowItem.PSObject.Properties) {
                        if($prop.Name -eq 'AdditionalFields') {
                            $row.Status = ($prop.Value.Field | Where-Object { $_.Title -eq "Status" }).Value
                            continue
                        }
                        
                        if ($properties.Contains($prop.Name)) {
                            $row.($prop.Name) = $prop.Value
                        }
                    }
                    $i++
                    [void]$Global:User.Add($row)  # Output immediately
                }

                # Update SearchStateContext for next page
                $sscNode = $response.SelectSingleNode("//core:SearchStateContext", $nsMgr)
                $searchStateContext = @{
                    Identifier   = $sscNode.Identifier
                    ResumeInfo   = $sscNode.ResumeInfo
                    CriteriaHash = $sscNode.CriteriaHash
                }

                # Continue if any values are non-empty
                $hasMore = ($searchStateContext.Identifier -or $searchStateContext.ResumeInfo -or $searchStateContext.CriteriaHash)
#break
            } while ($hasMore)

            $Global:UsersCacheTime = Get-Date
        }
        
        $Global:User
}

function Idm-UserGroupsRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'UserGroup'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return
        }

        # Refresh cache if needed
        if ($Global:Users.Count -eq 0) {
            Idm-UsersRead -SystemParams $SystemParams -FunctionParams $FunctionParams | Out-Null
        }

        # Precompute property template
        $properties = $Global:Properties.$Class | Where-Object { ('hidden' -notin $_.options ) }
        $propertiesHT = @{}; $Global:Properties.$Class | ForEach-Object { $propertiesHT[$_.name] = $_ }

        $template = [ordered]@{}
        foreach ($prop in $properties.Name) {
            $template[$prop] = $null
        }
        
        # Prepare runspace pool
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, [int]$system_params.nr_of_threads)
        $runspacePool.Open()
        $runspaces = @()

        # Index for tracking
        $index = 0

        $funcDef = "function Execute-Request { $((Get-Command Execute-Request -CommandType Function).ScriptBlock.ToString()) }"

        foreach ($item in $Global:User) {
            $runspace = [powershell]::Create().AddScript($funcDef).AddScript({
                param($item, $system_params, $Class, $template, $index)
                
                $itemResult = @{
                    rows = @()
                    logMessage = $null
                }

                $splat = @{
                    SystemParams = $system_params
                    Method = "POST"
                    Uri = '/interconnect-amcurprd-username/api/epic/2016/Security/PersonnelManagement/ViewUserGroups/Personnel/User/Groups/View'
                    Body = (@{ "UserID" = @{ "ID" = $item.ID; "Type"= "EXTERNAL"}}) | ConvertTo-Json
                    Class = $Class
                    LogMessage = "[$($item.ID)]"
                    LoggingEnabled = $false
                }

                try {
                    $response = Execute-Request @splat
                } catch {
                    $itemResult.logMessage = "Retrieve User Groups [$($item.ID)] - $_"
                    return $itemResult
                }

                foreach ($membershipRow in $response.UserGroups) {
                    $row = New-Object -TypeName PSObject -Property $template
                    $row.UserID = $item.Id
                    $row.Group = $membershipRow
                    $itemResult.rows += $row
                }

                return $itemResult
            }).AddArgument($item).AddArgument($system_params).AddArgument($Class).AddArgument($template).AddArgument($index)

            $runspace.RunspacePool = $runspacePool
            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke(); Index = $index }
            $index++
        }

        # Collect results
        $total = $runspaces.Count
        $completed = 0

        $result = @()
        foreach ($r in $runspaces) {
            $output = $r.Pipe.EndInvoke($r.Status)
            $completed++

            if ($completed % 250 -eq 0 -or $completed -eq $total) {
                $percent = [math]::Round(($completed / $total) * 100, 2)
                Log info "Progress: [$completed/$total] requests completed ($percent%)"
            }

            if($null -ne $output.logMessage) {
                Log verbose $output.logMessage
            }

            $result += $output.rows

            $r.Pipe.Dispose()
        }

        $runspacePool.Close()
        $runspacePool.Dispose()

        # Final output
        $result

}

function Idm-UserInfoRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'UserInfo'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return
        }

        # Refresh cache if needed
        if ($Global:Users.Count -eq 0) {
            Idm-UsersRead -SystemParams $SystemParams -FunctionParams $FunctionParams | Out-Null
        }

        $Global:UserInfo.Clear()
        $Global:UserInfo_UserIDs.Clear()

        # Precompute property template
        $properties = $Global:Properties.$Class | Where-Object { ('hidden' -notin $_.options ) }
        $propertiesHT = @{}; $Global:Properties.$Class | ForEach-Object { $propertiesHT[$_.name] = $_ }

        $template = [ordered]@{}
        foreach ($prop in $properties.Name) {
            if($propertiesHT[$prop].Type -eq 'object') {
                foreach($subProperty in $propertiesHT[$prop].objectfields) {
                    $template["$($prop)_$($subProperty)"] = $null
                }
                continue
            }

            $template[$prop] = $null
        }
        
        # Prepare runspace pool
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, [int]$system_params.nr_of_threads)
        $runspacePool.Open()
        $runspaces = @()

        # Index for tracking
        $index = 0

        $funcDef = "function Execute-Request { $((Get-Command Execute-Request -CommandType Function).ScriptBlock.ToString()) }"

        foreach ($item in $Global:User) {
            $runspace = [powershell]::Create().AddScript($funcDef).AddScript({
                param($item, $system_params, $Class, $template, $index, $properties, $propertiesHT)
                
                $itemResult = @{
                    result = $null
                    subResultUserIDs = [System.Collections.ArrayList]@()
                    subResultUserRoleIDs = [System.Collections.ArrayList]@()
                    subResultLinkedTemplatesConfig = [System.Collections.ArrayList]@()
                    subResultUserSubtemplateIDs = [System.Collections.ArrayList]@()
                    subResultLinkedProviderID = [System.Collections.ArrayList]@()
                    logMessage = $null
                }

                $splat = @{
                    SystemParams = $system_params
                    Uri = "/interconnect-amcurprd-username/api/epic/2014/Security/PersonnelManagement/ViewUser?UserID=$($item.ID)&UserType=EXTERNAL"
                    Body = $null
                    Class = $Class
                    LogMessage = "[$($item.ID)]"
                    LoggingEnabled = $false
                }

                try {
                    $response = Execute-Request @splat
                } catch {
                    $itemResult.logMessage = "Retrieve User Info [$($item.ID)] - $_"
                    return $itemResult
                }

                $row = New-Object -TypeName PSObject -Property ([ordered]@{} + $template)
                $row.UserID = $item.ID

                foreach ($prop in $response.PSObject.Properties) {     
                    if($prop.Name -eq 'UserIDs') {
                        foreach($record in $prop.Value) {
                            [void]$itemResult.subResultUserIDs.Add([PSCustomObject]@{ "UserID" = $row.UserID; "ID" = $record.ID; "Type" = $record.Type; })
                        }
                        continue
                    }

                    if($prop.Name -eq 'UserRoleIDs') {
                        foreach($record in $prop.Value) {
                            foreach($subRecord in $record.Identifiers) {
                                [void]$itemResult.subResultUserRoleIDs.Add([PSCustomObject]@{ "UserID" = $row.UserID; "ID" = $subRecord.ID; "Type" = $subRecord.Type; })
                            }
                        }
                        continue
                    }

                    if($prop.Name -eq 'LinkedTemplatesConfig') {
                        foreach($record in $prop.Value) {
                            foreach($subRecord in $record.DefaultTemplateID) {
                                [void]$itemResult.subResultLinkedTemplatesConfig.Add([PSCustomObject]@{ "UserID" = $row.UserID; "ID" = $subRecord.ID; "Type" = $subRecord.Type; "Area" = "DefaultTemplateID"})
                            }

                            foreach($subRecord in $record.AppliedTemplateID) {
                                [void]$itemResult.subResultLinkedTemplatesConfig.Add([PSCustomObject]@{ "UserID" = $row.UserID; "ID" = $subRecord.ID; "Type" = $subRecord.Type; "Area" = "AppliedTemplateID"})
                            }

                            foreach($subRecord in $record.AvailableLinkableTemplates) {
                                foreach($nestedSubRecord in $subRecord.LinkedTemplateID) {
                                    [void]$itemResult.subResultLinkedTemplatesConfig.Add([PSCustomObject]@{ "UserID" = $row.UserID; "ID" = $nestedSubRecord.ID; "Type" = $nestedSubRecord.Type; "Area" = "AvailableLinkableTemplates"})
                                }
                            }
                        }
                        continue
                    }

                    if($prop.Name -eq 'UserSubtemplateIDs') {
                        foreach($record in $prop.Value) {
                            foreach($subRecord in $record.Identifiers) {
                                [void]$itemResult.subResultUserSubtemplateIDs.Add([PSCustomObject]@{ "UserID" = $row.UserID; "ID" = $subRecord.ID; "Type" = $subRecord.Type; })
                            }
                        }
                        continue
                    }

                    if($prop.Name -eq 'LinkedProviderID') {
                        foreach($record in $prop.Value) {
                                [void]$itemResult.subResultLinkedProviderID.Add([PSCustomObject]@{ "UserID" = $row.UserID; "ID" = $record.ID; "Type" = $record.Type; })
                        }
                        continue
                    }

                    if($propertiesHT[$prop.Name].Type -eq 'object') {
                        foreach($subProperty in $propertiesHT[$prop.Name].objectfields) {
                            $row.("$($prop.Name)_$($subProperty)") = $prop.Value.$subProperty
                        }
                        continue
                    }
                    
                    if ($properties.Contains($prop.Name)) {
                        $row.($prop.Name) = $prop.Value
                    }
                }
                
                $itemResult.result = $row

                return $itemResult
            }).AddArgument($item).AddArgument($system_params).AddArgument($Class).AddArgument($template).AddArgument($index).AddArgument($properties).AddArgument($propertiesHT)

            $runspace.RunspacePool = $runspacePool
            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke(); Index = $index }
            $index++
        }

        # Collect results
        $total = $runspaces.Count
        $completed = 0

        foreach ($r in $runspaces) {
            $output = $r.Pipe.EndInvoke($r.Status)
            $completed++

            if ($completed % 250 -eq 0 -or $completed -eq $total) {
                $percent = [math]::Round(($completed / $total) * 100, 2)
                Log info "Progress: [$completed/$total] requests completed ($percent%)"
            }

            if($null -ne $output.logMessage) {
                Log verbose $output.logMessage
            }
            
            $output.result
            [void]$Global:UserInfo.AddRange(@() + $output.result)
            [void]$Global:UserInfo_UserIDs.Add($output.subResultUserIDs)
            [void]$Global:UserInfo_UserRoleIDs.Add($output.subResultUserRoleIDs)
            [void]$Global:UserInfo_LinkedTemplatesConfig.Add($output.subResultLinkedTemplatesConfig)
            [void]$Global:UserInfo_UserSubtemplateIDs.Add($output.subResultUserSubtemplateIDs)
            [void]$Global:UserInfo_LinkedProviderID.Add($output.subResultLinkedProviderID)

            $r.Pipe.Dispose()
        }

        $runspacePool.Close()
        $runspacePool.Dispose()

        # Final output
        $Global:UserInfoCacheTime = Get-Date
}

function Idm-UserInfo_UserIDsRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'UserInfo_UserID'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return
        }

        # Refresh cache if needed
        if ($Global:UserInfo.Count -eq 0) {
            Idm-UserInfoRead -SystemParams $SystemParams -FunctionParams $FunctionParams | Out-Null
        }

        $Global:UserInfo_UserIDs
}

function Idm-UserInfo_UserRoleIDsRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'UserInfo_UserID'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return
        }

        # Refresh cache if needed
        if ($Global:UserInfo.Count -eq 0) {
            Idm-UserInfoRead -SystemParams $SystemParams -FunctionParams $FunctionParams | Out-Null
        }

        $Global:UserInfo_UserRoleIDs
}

function Idm-UserInfo_LinkedTemplatesConfigRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'UserInfo_LinkedTemplatesConfig'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return
        }

        # Refresh cache if needed
        if ($Global:UserInfo.Count -eq 0) {
            Idm-UserInfoRead -SystemParams $SystemParams -FunctionParams $FunctionParams | Out-Null
        }

        $Global:UserInfo_LinkedTemplatesConfig
}

function Idm-UserInfo_UserSubtemplateIDsRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'UserInfo_UserSubtemplateIDs'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return
        }

        # Refresh cache if needed
        if ($Global:UserInfo.Count -eq 0) {
            Idm-UserInfoRead -SystemParams $SystemParams -FunctionParams $FunctionParams | Out-Null
        }

        $Global:UserInfo_UserSubtemplateIDs
}

function Idm-UserInfo_LinkedProviderIDRead {
    param (
        # Mode
        [switch] $GetMeta,    
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams

    )
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams
        $Class = 'UserInfo_LinkedProviderID'
        
        if ($GetMeta) {
            Get-ClassMetaData -SystemParams $SystemParams -Class $Class
            return
        }

        # Refresh cache if needed
        if ($Global:UserInfo.Count -eq 0) {
            Idm-UserInfoRead -SystemParams $SystemParams -FunctionParams $FunctionParams | Out-Null
        }

        $Global:UserInfo_LinkedProviderID
}

#
#   Internal Functions
#
function Get-ClassProperties {
    param (
        [string] $Class,
        [boolean] $IncludeHidden = $false
    )
    $properties = [System.Collections.ArrayList]@()

    foreach($prop in $Global:Properties.$Class) {
        if( ('hidden' -in $prop.options ) -and $IncludeHidden -eq $false ) { continue }
        if($prop.type -eq 'object') {
            if($prop.objectfields.count -gt 0) {
                [void]$properties.Add("$($prop.Name) { $($prop.objectfields -join ' ') }")
            }
            continue
        }

        [void]$properties.Add($prop.Name)
    }
    
    return $properties
}

function Execute-SOAPRequest {
    param (
        [hashtable] $SystemParams,
        [string] $Uri,
        [string] $Action,
        [string] $ActionResponse,
        [string] $Body,
        [string] $Class,
        [string] $LogMessage
    )
    # Generate Unique GUID for request
    $Guid = [guid]::NewGuid()
    
    # Generate CreatedUTC (current UTC time)
    $CreatedUTC = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    # Generate ExpiresUTC (5 minutes later)
    $ExpiresUTC = (Get-Date).ToUniversalTime().AddMinutes(5).ToString("yyyy-MM-ddTHH:mm:ssZ")

    #Build Request
    $soapBody = '<?xml version="1.0" encoding="utf-8"?>
        <s:Envelope
        xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
        xmlns:wsa="http://schemas.xmlsoap.org/ws/2003/03/addressing"
        xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"
        xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
            <s:Header>
                <wsa:MessageID>uuid:{0}</wsa:MessageID>
                <wsa:To s:mustUnderstand="1">https://vendorservices.epic.com/interconnect-amcurprd-personnel-username/httplistener.ashx</wsa:To>
                <wsa:Action s:mustUnderstand="1">{1}</wsa:Action>
                <wsse:Security s:mustUnderstand="1">
                    <wsu:Timestamp wsu:Id="TS-1">
                        <wsu:Created>{2}</wsu:Created>
                        <wsu:Expires>{3}</wsu:Expires>
                    </wsu:Timestamp>
                    <wsse:UsernameToken wsu:Id="UT-1">
                        <wsse:Username>emp:{4}</wsse:Username>
                        <wsse:Password
                Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">{5}</wsse:Password>
                    </wsse:UsernameToken>
                </wsse:Security>
            </s:Header>
            <s:Body>
                {6}
            </s:Body>
        </s:Envelope>' -f $Guid, $ActionResponse, $CreatedUTC, $ExpiresUTC, $SystemParams.username, $SystemParams.password, $Body

    # Build request parameters
    $splat = @{
        Headers = @{
            'Epic-Client-ID' = $SystemParams.clientId
            'SOAPAction' = $Action
            'Content-Type' = "text/xml; charset=utf-8"
        }
        Method = "POST"
        Uri = ("https://{0}{1}" -f $SystemParams.hostname, $uri)
        Body = $soapBody
    }

    # Configure proxy if enabled
    if ($SystemParams.use_proxy) {
        # Avoid redefining TrustAllCertsPolicy class unnecessarily
        #if (-not ([System.Net.ServicePointManager]::CertificatePolicy -is [TrustAllCertsPolicy])) {
            Add-Type -TypeDefinition @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@ -ErrorAction Stop
           # [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
       # }

        $splat["Proxy"] = $SystemParams.proxy_address

        if ($SystemParams.use_proxy_credentials) {
            $splat["ProxyCredential"] = New-Object System.Management.Automation.PSCredential (
                $SystemParams.proxy_username,
                (ConvertTo-SecureString $SystemParams.proxy_password -AsPlainText -Force)
            )
        }
    }

    $attempt = 0
    $retryDelay = $SystemParams.retryDelay

    do {
        try {
            $attemptSuffix = if ($attempt -gt 0) { " (Attempt $($attempt + 1))" } else { "" }
                    
            $baseMessage = "$($splat.Method) Call: $($splat.Uri)$($attemptSuffix) [$($Class)] $($LogMessage)"
            Log verbose $baseMessage

            return [xml](Invoke-WebRequest @splat -ErrorAction Stop).content
            break
        }
        catch {                
            $statusCode = $_.Exception.Response.StatusCode.value__
            if ($statusCode -eq 429 -and $attempt -lt $SystemParams.nr_of_retries) {
                $attempt++
                Log warning "Received $statusCode. Retrying in $retryDelay seconds..."
                Start-Sleep -Seconds $retryDelay
                $retryDelay *= 2 # Exponential backoff
            }
            else {
                throw "Failed request: $_"
            }
        }
    } while ($true)
}

function Execute-Request {
    param (
        [hashtable] $SystemParams,
        [string] $Method = "GET",
        [string] $Uri,
        [string] $Body,
        [string] $Class,
        [string] $LogMessage,
        [boolean] $LoggingEnabled = $true
    )
    # Setup  authorization token
    $pair = "emp`$$($SystemParams.username):$($SystemParams.password)"
    $encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))

    # Build request parameters
    $splat = @{
        Headers = @{
            Authorization = "Basic $($encodedCreds)"
            'Epic-Client-ID' = $SystemParams.clientId
            Accept = "application/json"
        }
        Method = $Method
        Uri = ("https://{0}{1}" -f $SystemParams.hostname, $uri)
    }

    if(@('PATCH','PUT','POST') -contains $method ) { 
        $splat.Body = $Body 
        $splat.Headers.'Content-Type' = "application/json"
    }

    # Configure proxy if enabled
    if ($SystemParams.use_proxy) {
        # Avoid redefining TrustAllCertsPolicy class unnecessarily
        if (-not ([System.Net.ServicePointManager]::CertificatePolicy -is [TrustAllCertsPolicy])) {
            Add-Type -TypeDefinition @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@ -ErrorAction Stop
            [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
        }

        $splat["Proxy"] = $SystemParams.proxy_address

        if ($SystemParams.use_proxy_credentials) {
            $splat["ProxyCredential"] = New-Object System.Management.Automation.PSCredential (
                $SystemParams.proxy_username,
                (ConvertTo-SecureString $SystemParams.proxy_password -AsPlainText -Force)
            )
        }
    }

    $attempt = 0
    $retryDelay = $SystemParams.retryDelay

    do {
        try {
            $attemptSuffix = if ($attempt -gt 0) { " (Attempt $($attempt + 1))" } else { "" }
            
            $baseMessage = "$($splat.Method) Call: $($splat.Uri)$($attemptSuffix) [$($Class)] $($LogMessage)"
            if($LoggingEnabled) { Log verbose $baseMessage }

            $response = Invoke-RestMethod @splat -ErrorAction Stop
            break
        }
        catch {
            if($errorBreak) { throw}
            
            $statusCode = $_.Exception.Response.StatusCode.value__
            if ($statusCode -eq 429 -and $attempt -lt $SystemParams.nr_of_retries) {
                $attempt++
                if($LoggingEnabled) { Log warning "Received $statusCode. Retrying in $retryDelay seconds..." }
                Start-Sleep -Seconds $retryDelay
                $retryDelay *= 2 # Exponential backoff
            }
            else {
                throw "Failed request: $_"
            }
        }
    } while ($true)

    return $response
}

function Get-ClassMetaData {
    param (
        [string] $SystemParams,
        [string] $Class
    )

    @(
        @{
            name = 'properties'
            type = 'grid'
            label = 'Properties'
            table = @{
                rows = @( $Global:Properties.$Class | Where-Object {'hidden' -notin $_.options } | ForEach-Object {
                    @{
                        name = $_.name
                        usage_hint = @( @(
                            foreach ($opt in $_.options) {
                                if ($opt -notin @('default', 'idm', 'key')) { continue }

                                if ($opt -eq 'idm') {
                                    $opt.Toupper()
                                }
                                else {
                                    $opt.Substring(0,1).Toupper() + $opt.Substring(1)
                                }
                            }
                        ) | Sort-Object) -join ' | '
                    }
                })
                settings_grid = @{
                    selection = 'multiple'
                    key_column = 'name'
                    checkbox = $true
                    filter = $true
                    columns = @(
                        @{
                            name = 'name'
                            display_name = 'Name'
                        }
                        @{
                            name = 'usage_hint'
                            display_name = 'Usage hint'
                        }
                    )
                }
            }
            value = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
    )
}
