# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("SharePoint Online") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> EntraIdAppId
$tmpName = @'
EntraIdAppId
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False" });

#Global variable #2 >> EntraIdTenantId
$tmpName = @'
EntraIdTenantId
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False" });

#Global variable #3 >> EntraIdCertificatePassword
$tmpName = @'
EntraIdCertificatePassword
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False" });

#Global variable #4 >> SharePointBaseUrl
$tmpName = @'
SharePointBaseUrl
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False" });

#Global variable #5 >> EntraIdCertificateBase64String
$tmpName = @'
EntraIdCertificateBase64String
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False" });


#make sure write-information logging is visual
$InformationPreference = "continue"

# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic }
    Write-Information "Using prefilled API credentials"
}
else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key }
    Write-Information "Using manual API credentials"
}

# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
}
else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}

# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  

# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if ($IsCoreCLR -eq $true) {
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return , $r  # Force return value to be an array using a comma
    }
    else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return , $r  # Force return value to be an array using a comma
    }
}

function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )

    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false

        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid

            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
        else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    }
    catch {
        Write-Error "Variable '$Name', message: $_"
    }
}

function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter { $_.name -eq $TaskName }

        if ([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task

            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid

            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
        else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    }
    catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }

    $returnObject.Value = $taskGuid
}

function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter()][String][AllowEmptyString()]$DatasourceRunInCloud,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    $datasourceTypeName = switch ($DatasourceType) { 
        "1" { "Native data source"; break } 
        "2" { "Static data source"; break } 
        "3" { "Task data source"; break } 
        "4" { "Powershell data source"; break }
    }

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if ([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
                runInCloud         = $DatasourceRunInCloud;
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl + "api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
        else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    }
    catch {
        Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }

    $returnObject.Value = $datasourceGuid
}

function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl + "api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        }
        catch {
            $response = $null
        }

        if (([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
        else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    }
    catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }

    $returnObject.Value = $formGuid
}


function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl + "api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        }
        catch {
            $response = $null
        }

        if ([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if (-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true

            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl + "api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        }
        else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    }
    catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }

    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}

<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
    Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "sharepoint-online-site-manage-members | Sharepoint-Get-Site-Group-Members" #>
$tmpPsScript = @'
$groupUrl = $datasource.selectedGroup.UsersUri

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}


function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Resource
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            #resource              = 'https://graph.microsoft.com'
            scope                 = "$resource/.default"
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/v2.0/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-SharePointGroupMembers {
    param(
        [string]$SharePointToken,
        [string]$TenantName     
    )

    $headers = @{
        "Authorization" = "Bearer $SharePointToken"
        "Accept"        = "application/json;odata=verbose"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $groupUrl -Method Get -Headers $headers
        return $response.d.results
    }
    catch {
        Write-Error "Failed to retrieve site groups: $_"
        throw
    }
}
#endregion functions

try {
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword
    Write-Verbose "Converted base64 certificate string to certificate object"

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId -Resource $SharePointBaseUrl 
    Write-Verbose "Created access token"

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"            
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }
    Write-Verbose "Created headers"

    $getMembersSplatParams = @{
        SharePointToken = $entraToken                                
    }

    $siteGroupMembers = Get-SharePointGroupMembers @getMembersSplatParams | Where-Object { $null -ne $_.UserPrincipalName } | Select-Object -Unique *

    foreach ($sitemember in $siteGroupMembers) {          
        $returnObject = @{displayName = $($sitemember.Title); userPrincipalName = $($sitemember.UserPrincipalName) }
        Write-Output $returnObject                
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}

'@ 
$tmpModel = @'
[{"key":"displayName","type":0},{"key":"userPrincipalName","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedSite","type":0,"options":1},{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedGroup","type":0,"options":1}]
'@ 
$dataSourceGuid_3 = [PSCustomObject]@{} 
$dataSourceGuid_3_Name = @'
sharepoint-online-site-manage-members | Sharepoint-Get-Site-Group-Members
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_3_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_3) 
<# End: DataSource "sharepoint-online-site-manage-members | Sharepoint-Get-Site-Group-Members" #>

<# Begin: DataSource "sharepoint-online-site-manage-members | Sharepoint-Get-Site-Groups" #>
$tmpPsScript = @'
$siteUrl = $datasource.selectedSite.Site

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}


function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Resource
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            #resource              = 'https://graph.microsoft.com'
            scope                 = "$resource/.default"
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/v2.0/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-SharePointGroups {
    param(
        [string]$SharePointToken,
        [string]$TenantName     
    )

    
    
    $apiUrl = "$siteUrl/_api/web/SiteGroups"
    
    $headers = @{
        "Authorization" = "Bearer $SharePointToken"
        "Accept"        = "application/json;odata=verbose"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
        return $response.d.results
    }
    catch {
        Write-Error "Failed to retrieve site groups: $_"
        throw
    }
}
#endregion functions

try {
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword
    Write-Verbose "Converted base64 certificate string to certificate object"

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId -Resource "https://jbt4e.sharepoint.com" 
    Write-Verbose "Created access token"

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"            
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }
    Write-Verbose "Created headers"

    $getSiteGroupsSplatParams = @{
        SharePointToken = $entraToken                                
    }
    $siteGroups = Get-SharePointGroups @getSiteGroupsSplatParams

    foreach ($siteGroup in $siteGroups) {                             
        $returnObject = [ordered]@{GroupName = $siteGroup.Title; OwnersUri = $siteGroup.Owner."__deferred".uri; UsersUri = $siteGroup.Users."__deferred".uri; Id=$siteGroup.Id }        
        Write-Output $returnObject                
    }
    
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
    

'@ 
$tmpModel = @'
[{"key":"GroupName","type":0},{"key":"OwnersUri","type":0},{"key":"UsersUri","type":0},{"key":"Id","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedSite","type":0,"options":1}]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
sharepoint-online-site-manage-members | Sharepoint-Get-Site-Groups
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "sharepoint-online-site-manage-members | Sharepoint-Get-Site-Groups" #>

<# Begin: DataSource "sharepoint-onlne-site-manage-members | EntraID-Get-All-Users" #>
$tmpPsScript = @'
# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
$filter = "`$filter=displayName ne null"
$propertiesToSelect = @("id", "displayName", "mail", "userPrincipalName", "employeeId", "accountEnabled") # Properties to select from Microsoft Graph API, comma separated

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}


function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions

try {
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword
    Write-Verbose "Converted base64 certificate string to certificate object"

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId
    Write-Verbose "Created access token"

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }
    Write-Verbose "Created headers"

    # Get Microsoft Entra ID Users
    # API docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying all Microsoft Entra ID Users"
    $microsoftEntraIDUsers = [System.Collections.ArrayList]@()
    do {
        $getMicrosoftEntraIDUsersSplatParams = @{
            Uri         = "https://graph.microsoft.com/v1.0/users?$filter&`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"
            Headers     = $headers
            Method      = "GET"
            Verbose     = $false
            ErrorAction = "Stop"
        }
        if (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDUsersResponse.'@odata.nextLink')) {
            $getMicrosoftEntraIDUsersSplatParams["Uri"] = $getMicrosoftEntraIDUsersResponse.'@odata.nextLink'
        }
        
        $getMicrosoftEntraIDUsersResponse = $null
        $getMicrosoftEntraIDUsersResponse = Invoke-RestMethod @getMicrosoftEntraIDUsersSplatParams
    
        # Select only specified properties to limit memory usage
        $getMicrosoftEntraIDUsersResponse.Value = $getMicrosoftEntraIDUsersResponse.Value | Select-Object $propertiesToSelect

        if ($getMicrosoftEntraIDUsersResponse.Value -is [array]) {
            [void]$microsoftEntraIDUsers.AddRange($getMicrosoftEntraIDUsersResponse.Value)
        }
        else {
            [void]$microsoftEntraIDUsers.Add($getMicrosoftEntraIDUsersResponse.Value)
        }
    } while (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDUsersResponse.'@odata.nextLink'))
    Write-Information "Queried all Microsoft Entra ID Users. Result count: $(($microsoftEntraIDUsers | Measure-Object).Count)"

    # Send results to HelloID
    $actionMessage = "sending results to HelloID"
    $microsoftEntraIDUsers | ForEach-Object {
        # Add displayValue property as HelloID can only display a single property
        $_ | Add-Member -NotePropertyName 'displayValue' -NotePropertyValue "$($_.displayName) [$($_.userPrincipalName)]" -Force
        Write-Output $_
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
'@ 
$tmpModel = @'
[{"key":"id","type":0},{"key":"displayName","type":0},{"key":"mail","type":0},{"key":"userPrincipalName","type":0},{"key":"employeeId","type":0},{"key":"accountEnabled","type":0},{"key":"displayValue","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_2 = [PSCustomObject]@{} 
$dataSourceGuid_2_Name = @'
sharepoint-onlne-site-manage-members | EntraID-Get-All-Users
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_2_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "False" -returnObject ([Ref]$dataSourceGuid_2) 
<# End: DataSource "sharepoint-onlne-site-manage-members | EntraID-Get-All-Users" #>

<# Begin: DataSource "sharepoint-online-site-manage-members | SharePoint-Get-Sites-Wildcard" #>
$tmpPsScript = @'
# Variables configured in form
$searchValue = $datasource.searchvalue
if ($searchValue -eq "*") {
    $filter = "`$filter=groupTypes/any(c:c+eq+'Unified')"
}
else {
    $filter = "`$search=`"displayName:$searchValue`" OR `"description:$searchValue`" OR `"mail:$searchValue`"&`$filter=groupTypes/any(c:c eq 'Unified')"
}

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}


function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-GroupSharePointSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $GroupId
    )
    
    try {
        $getSiteSplatParams = @{
            Uri = "https://graph.microsoft.com/v1.0/groups/$groupId/sites/root"
            Headers = $headers
            Method = 'GET'
            Verbose     = $false
            ErrorAction = "Stop"
        }
        
        $site = Invoke-RestMethod @getSiteSplatParams
        return $site
    }
    catch {
        Write-Information $_.Exception.Message
        return $null
    }
}
#endregion functions

try {
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword
    Write-Verbose "Converted base64 certificate string to certificate object"

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId
    Write-Verbose "Created access token"

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }
    Write-Verbose "Created headers"

    # Get Microsoft Entra ID Groups
    # API docs: https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Groups matching search value [$searchValue]"
    $microsoftEntraIDGroups = [System.Collections.ArrayList]@()
    do {
        $getMicrosoftEntraIDGroupsSplatParams = @{
            Uri         = "https://graph.microsoft.com/v1.0/groups?$filter&`$top=999&`$count=true"
            Headers     = $headers
            Method      = "GET"
            Verbose     = $false
            ErrorAction = "Stop"
        }
        if (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDGroupsResponse.'@odata.nextLink')) {
            $getMicrosoftEntraIDGroupsSplatParams["Uri"] = $getMicrosoftEntraIDGroupsResponse.'@odata.nextLink'
        }
        
        $getMicrosoftEntraIDGroupsResponse = $null
        $getMicrosoftEntraIDGroupsResponse = Invoke-RestMethod @getMicrosoftEntraIDGroupsSplatParams
    
        # Select only specified properties to limit memory usage
        $getMicrosoftEntraIDGroupsResponse.Value = $getMicrosoftEntraIDGroupsResponse.Value | Select-Object $propertiesToSelect

        if ($getMicrosoftEntraIDGroupsResponse.Value -is [array]) {
            [void]$microsoftEntraIDGroups.AddRange($getMicrosoftEntraIDGroupsResponse.Value)
        }
        else {
            [void]$microsoftEntraIDGroups.Add($getMicrosoftEntraIDGroupsResponse.Value)
        }
    } while (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDGroupsResponse.'@odata.nextLink'))
    Write-Information "Queried Microsoft Entra ID Groups matching search value [$searchValue]. Result count: $(($microsoftEntraIDGroups | Measure-Object).Count)"

        $resultCount = @($microsoftEntraIDGroups ).Count
        Write-Information -Message "Result count: $resultCount"
         
        if($resultCount -gt 0){
            foreach($group in $microsoftEntraIDGroups){

                $getGroupSplatParams = @{
                    GroupId = $group.Id
                }
                
                $site = Get-GroupSharePointSite @getGroupSplatParams 
                
                if($site){               
                    $returnObject = @{DisplayName=$group.DisplayName; Description=$group.Description; MailNickName=$group.MailNickName; GroupId=$group.Id; Site=$site.WebUrl; SiteId=$site.id}
                    Write-Output $returnObject
                }
            }
        } else {
            return
        }
    
} catch {
     $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}

'@ 
$tmpModel = @'
[{"key":"Site","type":0},{"key":"Description","type":0},{"key":"DisplayName","type":0},{"key":"GroupId","type":0},{"key":"SiteId","type":0},{"key":"MailNickName","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"searchValue","type":0,"options":1}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
sharepoint-online-site-manage-members | SharePoint-Get-Sites-Wildcard
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "sharepoint-online-site-manage-members | SharePoint-Get-Sites-Wildcard" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "SharePoint online site - manage members" #>
$tmpSchema = @"
[{"label":"Select Site","fields":[{"key":"searchValue","templateOptions":{"label":"Search for SiteName","required":true},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"sites","templateOptions":{"label":"Select Site","required":true,"grid":{"columns":[{"headerName":"Display Name","field":"DisplayName"},{"headerName":"Description","field":"Description"},{"headerName":"Site","field":"Site"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[{"propertyName":"searchValue","otherFieldValue":{"otherFieldKey":"searchValue"}}]}},"useDefault":false,"useFilter":true,"allowCsvDownload":true},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]},{"label":"Site groups","fields":[{"key":"sharepointGroups","templateOptions":{"label":"Select Group","required":true,"grid":{"columns":[{"headerName":"Group Name","field":"GroupName"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[{"propertyName":"selectedSite","otherFieldValue":{"otherFieldKey":"sites"}}]}},"useDefault":false,"allowCsvDownload":true},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]},{"label":"Manage Members","fields":[{"key":"members","templateOptions":{"label":"Manage members","required":false,"filterable":false,"useDataSource":true,"dualList":{"options":[{"guid":"75ea2890-88f8-4851-b202-626123054e14","Name":"Apple"},{"guid":"0607270d-83e2-4574-9894-0b70011b663f","Name":"Pear"},{"guid":"1ef6fe01-3095-4614-a6db-7c8cd416ae3b","Name":"Orange"}],"optionKeyProperty":"userPrincipalName","optionDisplayProperty":"displayName","labelLeft":"Available","labelRight":"Current"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_2","input":{"propertyInputs":[]}},"destinationDataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_3","input":{"propertyInputs":[{"propertyName":"selectedSite","otherFieldValue":{"otherFieldKey":"sites"}},{"propertyName":"selectedGroup","otherFieldValue":{"otherFieldKey":"sharepointGroups"}}]}}},"type":"duallist","summaryVisibility":"Show","sourceDataSourceIdentifierSuffix":"source-datasource","destinationDataSourceIdentifierSuffix":"destination-datasource","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
SharePoint online site - manage members
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if (-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)) {
    foreach ($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl + "api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
        
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        }
        catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if ($null -ne $delegatedFormAccessGroupGuids) {
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}

$delegatedFormCategoryGuids = @()
foreach ($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl + "api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $response = $response | Where-Object { $_.name.en -eq $category }
    
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
    
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
    catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category };
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100

        $uri = ($script:PortalBaseUrl + "api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid

        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null } 
$delegatedFormName = @'
SharePoint online site - manage members
'@
$tmpTask = @'
{"name":"SharePoint online site - manage members","script":"# variables configured in form:\n$groupId = $form.sharepointGroups.GroupName\n$membersToAdd = $form.members.leftToRight | Select-Object -Unique *\n$membersToRemove = $form.members.rightToLeft | Select-Object -Unique *\n$siteUrl = $form.sites.Site\n$groupId = $form.sharepointGroups.Id\n\n# Global variables\n# Outcommented as these are set from Global Variables\n# $EntraIdTenantId = \"\"\n# $EntraIdAppId = \"\"\n# $EntraIdCertificateBase64String = \"\"\n# $EntraIdCertificatePassword = \"\"\n\n# Set TLS to accept TLS, TLS 1.1 and TLS 1.2\n[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12\n\n# Set debug logging\n$VerbosePreference = \"SilentlyContinue\"\n$InformationPreference = \"Continue\"\n$WarningPreference = \"Continue\"\n\n#region functions\nfunction Resolve-MicrosoftGraphAPIError {\n    [CmdletBinding()]\n    param (\n        [Parameter(Mandatory)]\n        [object]\n        $ErrorObject\n    )\n    process {\n        $httpErrorObj = [PSCustomObject]@{\n            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber\n            Line             = $ErrorObject.InvocationInfo.Line\n            ErrorDetails     = $ErrorObject.Exception.Message\n            FriendlyMessage  = $ErrorObject.Exception.Message\n        }\n        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {\n            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message\n        }\n        elseif ($ErrorObject.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027) {\n            if ($null -ne $ErrorObject.Exception.Response) {\n                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()\n                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {\n                    $httpErrorObj.ErrorDetails = $streamReaderResponse\n                }\n            }\n        }\n        try {\n            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)\n            if ($errorDetailsObject.error_description) {\n                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description\n            }\n            elseif ($errorDetailsObject.error.message) {\n                $httpErrorObj.FriendlyMessage = \"$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)\"\n            }\n            elseif ($errorDetailsObject.error.details.message) {\n                $httpErrorObj.FriendlyMessage = \"$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)\"\n            }\n            else {\n                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails\n            }\n        }\n        catch {\n            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails\n        }\n        Write-Output $httpErrorObj\n    }\n}\n\n\nfunction Get-MSEntraAccessToken {\n    [CmdletBinding()]\n    param(\n        [Parameter(Mandatory)]\n        [ValidateNotNull()]\n        $Certificate,\n        \n        [Parameter(Mandatory)]\n        [ValidateNotNullOrEmpty()]\n        [string]\n        $AppId,\n        \n        [Parameter(Mandatory)]\n        [ValidateNotNullOrEmpty()]\n        [string]\n        $TenantId,\n        \n        [Parameter(Mandatory)]\n        [ValidateNotNullOrEmpty()]\n        [string]\n        $Resource\n    )\n    try {\n        # Get the DER encoded bytes of the certificate\n        $derBytes = $Certificate.RawData\n\n        # Compute the SHA-256 hash of the DER encoded bytes\n        $sha256 = [System.Security.Cryptography.SHA256]::Create()\n        $hashBytes = $sha256.ComputeHash($derBytes)\n        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\n\n        # Create a JWT (JSON Web Token) header\n        $header = @{\n            \u0027alg\u0027      = \u0027RS256\u0027\n            \u0027typ\u0027      = \u0027JWT\u0027\n            \u0027x5t#S256\u0027 = $base64Thumbprint\n        } | ConvertTo-Json\n        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))\n\n        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for \u0027exp\u0027, \u0027nbf\u0027 and \u0027iat\u0027\n        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]\u00271970-01-01T00:00:00Z\u0027).ToUniversalTime()).TotalSeconds)\n\n        # Create a JWT payload\n        $payload = [Ordered]@{\n            \u0027iss\u0027 = \"$($AppId)\"\n            \u0027sub\u0027 = \"$($AppId)\"\n            \u0027aud\u0027 = \"https://login.microsoftonline.com/$($TenantId)/oauth2/token\"\n            \u0027exp\u0027 = ($currentUnixTimestamp + 3600) # Expires in 1 hour\n            \u0027nbf\u0027 = ($currentUnixTimestamp - 300) # Not before 5 minutes ago\n            \u0027iat\u0027 = $currentUnixTimestamp\n            \u0027jti\u0027 = [Guid]::NewGuid().ToString()\n        } | ConvertTo-Json\n        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\n\n        # Extract the private key from the certificate\n        $rsaPrivate = $Certificate.PrivateKey\n        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()\n        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))\n\n        # Sign the JWT\n        $signatureInput = \"$base64Header.$base64Payload\"\n        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), \u0027SHA256\u0027)\n        $base64Signature = [System.Convert]::ToBase64String($signature).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\n\n        # Create the JWT token\n        $jwtToken = \"$($base64Header).$($base64Payload).$($base64Signature)\"\n\n        $createEntraAccessTokenBody = @{\n            grant_type            = \u0027client_credentials\u0027\n            client_id             = $AppId\n            client_assertion_type = \u0027urn:ietf:params:oauth:client-assertion-type:jwt-bearer\u0027\n            client_assertion      = $jwtToken\n            #resource              = \u0027https://graph.microsoft.com\u0027\n            scope                 = \"$resource/.default\"\n        }\n\n        $createEntraAccessTokenSplatParams = @{\n            Uri         = \"https://login.microsoftonline.com/$($TenantId)/oauth2/v2.0/token\"\n            Body        = $createEntraAccessTokenBody\n            Method      = \u0027POST\u0027\n            ContentType = \u0027application/x-www-form-urlencoded\u0027\n            Verbose     = $false\n            ErrorAction = \u0027Stop\u0027\n        }\n\n        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams\n        Write-Output $createEntraAccessTokenResponse.access_token\n    }\n    catch {\n        $PSCmdlet.ThrowTerminatingError($_)\n    }\n}\n\nfunction Get-MSEntraCertificate {\n    [CmdletBinding()]\n    param(\n        [Parameter(Mandatory)]\n        [ValidateNotNullOrEmpty()]\n        [string]\n        $CertificateBase64String,\n        \n        [Parameter(Mandatory)]\n        [ValidateNotNullOrEmpty()]\n        [string]\n        $CertificatePassword\n    )\n    try {\n        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)\n        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)\n        Write-Output $certificate\n    }\n    catch {\n        $PSCmdlet.ThrowTerminatingError($_)\n    }\n}\n\nfunction Add-SharePointGroupMembers {\n    param(\n        [string]$SharePointToken,\n        [string]$TenantName,\n        [string]$userName     \n    )\n\n    $headers = @{\n        \"Authorization\" = \"Bearer $SharePointToken\"\n        \"Accept\"        = \"application/json\"\n    }\n    \n    $body = @\"\n    {\n        \"__metadata\": { \"type\": \"SP.User\" },\n        \"LoginName\":\"i:0#.f|membership|$userName\"        \n    }\n\"@\n\n    $addMemberToGroupSplatParams = @{\n        Uri = \"$siteUrl/_api/Web/SiteGroups($groupId)/Users\"\n        Method = \u0027POST\u0027\n        Headers = $headers\n        Body = $body\n        ContentType = \"application/json;odata=verbose\"\n    }\n\n    try {\n        $response = Invoke-RestMethod @addMemberToGroupSplatParams        \n    }\n    catch {\n        Write-Error \"Failed to add member to group: $_\"\n        throw\n    }\n}\n\nfunction Remove-SharePointGroupMembers {\n    param(\n        [string]$SharePointToken,\n        [string]$TenantName,        \n        [string]$userId\n    )\n\n    $headers = @{\n        \"Authorization\" = \"Bearer $SharePointToken\"\n        \"Accept\"        = \"application/json\"\n    }\n\n    $removeMemberToGroupSplatParams = @{\n        Uri = \"$siteUrl/_api/Web/SiteGroups($groupId)/Users/RemoveByID($userId)\"\n        Method = \u0027POST\u0027\n        Headers = $headers\n        ContentType = \"application/json;odata=verbose\"\n    }\n    \n    try {\n        $response = Invoke-RestMethod @removeMemberToGroupSplatParams        \n    }\n    catch {\n        Write-Error \"Failed to remove member from group: $_\"\n        throw\n    }\n}\n\nfunction Get-SharePointUserIdByUsername {\n    param(\n        [string]$SharePointToken,\n        [string]$TenantName,\n        [string]$userName     \n    )\n\n    $headers = @{\n        \"Authorization\" = \"Bearer $SharePointToken\"\n        \"Accept\"        = \"application/json\"\n    }\n    \n    $getUserIdSplatParams = @{\n        #_api/web/SiteGroups/GetById(3)/Users?$filter=Email eq \u0027UserEmail@email.com\u0027\n        Uri = \"$siteUrl/_api/Web/SiteGroups/GetById($groupId)/Users\"\n        Method = \u0027GET\u0027\n        Headers = $headers\n    }\n\n    try {\n        $response = Invoke-RestMethod @getUserIdSplatParams           \n        return ($response.value | Where-Object { $_.UserPrincipalName -eq $userName}).ID\n    }\n    catch {\n        Write-Error \"Failed to retrieve userid of member: $_\"\n        throw\n    }\n}\n#endregion functions\n\ntry {\n    # Convert base64 certificate string to certificate object\n    $actionMessage = \"converting base64 certificate string to certificate object\"\n    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword\n    Write-Verbose \"Converted base64 certificate string to certificate object\"\n\n    # Create access token\n    $actionMessage = \"creating access token\"\n    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId -Resource $SharePointBaseUrl \n    Write-Verbose \"Created access token\"\n\n    # Create headers\n    $actionMessage = \"creating headers\"\n    $headers = @{\n        \"Authorization\"    = \"Bearer $($entraToken)\"            \n        \"Accept\"           = \"application/json\"\n        \"Content-Type\"     = \"application/json\"\n        \"ConsistencyLevel\" = \"eventual\" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)\n    }\n    Write-Verbose \"Created headers\"\n\n    $actionMessage = \"add members to group\"\n    foreach ($user in $membersToAdd) {        \n         $addMembersSplatParams = @{\n            SharePointToken = $entraToken \n            TenantName = $SharePointBaseUrl        \n            UserName = $user.userPrincipalName                       \n        }        \n        $null = Add-SharePointGroupMembers @addMembersSplatParams\n\n        Write-Information \"Successfully added User [$($user.displayName)] to Members of [$($form.sharepointGroups.GroupName)]\"\n        $Log = @{\n            Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \n            System            = \"SharePoint\" # optional (free format text) \n            Message           = \"Successfully added User [$($user.displayName)] to Members of [$($form.sharepointGroups.GroupName)]\" # required (free format text) \n            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n            TargetDisplayName = $($user.displayName) # optional (free format text) \n            TargetIdentifier  = $($form.sites.GroupId) # optional (free format text) \n        }\n        #send result back  \n        Write-Information -Tags \"Audit\" -MessageData $log\n    }\n\n    $actionMessage = \"remove members from group\"\n    foreach ($user in $membersToRemove) {\n        $getUserIdSplatParams = @{\n            SharePointToken = $entraToken \n            TenantName = $SharePointBaseUrl        \n            UserName = $user.UserPrincipalName                       \n        }\n\n        $userId = Get-SharePointUserIdByUsername @getUserIdSplatParams\n\n        $removeMembersSplatParams = @{\n            SharePointToken = $entraToken \n            TenantName = $SharePointBaseUrl        \n            UserId = $userid\n        }        \n        $null = Remove-SharePointGroupMembers @removeMembersSplatParams    \n        \n        Write-Information \"Successfully removed User [$($user.DisplayName)] from Members of [$($form.sharepointGroups.GroupName)]\"\n        $Log = @{\n            Action            = \"RevokeMembership\" # optional. ENUM (undefined = default) \n            System            = \"SharePoint\" # optional (free format text) \n            Message           = \"Successfully removed User [$($user.DisplayName)] from Members of [$($form.sharepointGroups.GroupName)]\" # required (free format text) \n            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n            TargetDisplayName = $($user.Name) # optional (free format text) \n            TargetIdentifier  = $($form.sites.GroupId) # optional (free format text) \n        }\n        #send result back  \n        Write-Information -Tags \"Audit\" -MessageData $log        \n    }       \n}\ncatch {\n     $ex = $PSItem\n    if ($($ex.Exception.GetType().FullName -eq \u0027Microsoft.PowerShell.Commands.HttpResponseException\u0027) -or\n        $($ex.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027)) {\n        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex\n        $auditMessage = \"Error $($actionMessage). Error: $($errorObj.FriendlyMessage)\"\n        $warningMessage = \"Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)\"\n    }\n    else {\n        $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\n        $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\n    }\n\n    $Log = @{\n        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \n        System            = \"SharePoint\" # optional (free format text) \n        Message           = $auditMessage # required (free format text) \n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $($form.sharepointGroups.GroupName) # optional (free format text) \n        TargetIdentifier  = $($form.sites.GroupId) # optional (free format text) \n    }\n    \n    Write-Information -Tags \"Audit\" -MessageData $log\n    Write-Warning $warningMessage\n    Write-Error $auditMessage\n}","runInCloud":false}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-puzzle-piece" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

