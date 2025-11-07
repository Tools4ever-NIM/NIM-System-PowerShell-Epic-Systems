# version: 1.0
#
# Vector Solutions SafeSchools.ps1 - Vector Solution SafeSchools
#
$Log_MaskableKeys = @(
    'password',
    "proxy_password",
    "clientsecret"
)

$Global:PeopleCacheTime = Get-Date
$Global:People = [System.Collections.ArrayList]@()

$Properties = @{
    User = @(
        @{ name = 'ID';              type = 'string';   objectfields = $null;             options = @('default','key') },
        @{ name = 'Name';            type = 'string'; objectfields = $null;             options = @('default') },
        @{ name = 'ExternalID';      type = 'string'; objectfields = $null;             options = @('default') },
        @{ name = 'EnRolRank';       type = 'string'; objectfields = $null;             options = @('default') },
        @{ name = 'Status';          type = 'string';   objectfields = @('personId');     options = @('default') }
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
                name = 'pageSize'
                type = 'textbox'
                label = 'Page Size'
                description = 'Size of each request page'
                value = 1000
            }
            @{
                name = 'retryDelay'
                type = 'textbox'
                label = 'Seconds to wait for retry'
                description = ''
                value = 2
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
        <INI>EMP</INI>
        <SearchString></SearchString>
        <RecordState>Active</RecordState>
        <SkipEnRol>false</SkipEnRol>
        <SoundsLikeMode>UseIfNeeded</SoundsLikeMode>
    </SearchCriteria>
    $searchStateXml
    <UserID>1</UserID>
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
                $row  # Output immediately
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
        } while ($hasMore)

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
