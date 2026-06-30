# variables configured in form:
$groupId = $form.sharepointGroups.GroupName
$membersToAdd = $form.members.leftToRight | Select-Object -Unique *
$membersToRemove = $form.members.rightToLeft | Select-Object -Unique *
$siteUrl = $form.sites.Site
$groupId = $form.sharepointGroups.Id

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

function Add-SharePointGroupMembers {
    param(
        [string]$SharePointToken,
        [string]$TenantName,
        [string]$userName     
    )

    $headers = @{
        "Authorization" = "Bearer $SharePointToken"
        "Accept"        = "application/json"
    }
    
    $body = @"
    {
        "__metadata": { "type": "SP.User" },
        "LoginName":"i:0#.f|membership|$userName"        
    }
"@

    $addMemberToGroupSplatParams = @{
        Uri = "$siteUrl/_api/Web/SiteGroups($groupId)/Users"
        Method = 'POST'
        Headers = $headers
        Body = $body
        ContentType = "application/json;odata=verbose"
    }

    try {
        $response = Invoke-RestMethod @addMemberToGroupSplatParams        
    }
    catch {
        Write-Error "Failed to add member to group: $_"
        throw
    }
}

function Remove-SharePointGroupMembers {
    param(
        [string]$SharePointToken,
        [string]$TenantName,        
        [string]$userId
    )

    $headers = @{
        "Authorization" = "Bearer $SharePointToken"
        "Accept"        = "application/json"
    }

    $removeMemberToGroupSplatParams = @{
        Uri = "$siteUrl/_api/Web/SiteGroups($groupId)/Users/RemoveByID($userId)"
        Method = 'POST'
        Headers = $headers
        ContentType = "application/json;odata=verbose"
    }
    
    try {
        $response = Invoke-RestMethod @removeMemberToGroupSplatParams        
    }
    catch {
        Write-Error "Failed to remove member from group: $_"
        throw
    }
}

function Get-SharePointUserIdByUsername {
    param(
        [string]$SharePointToken,
        [string]$TenantName,
        [string]$userName     
    )

    $headers = @{
        "Authorization" = "Bearer $SharePointToken"
        "Accept"        = "application/json"
    }
    
    $getUserIdSplatParams = @{
        #_api/web/SiteGroups/GetById(3)/Users?$filter=Email eq 'UserEmail@email.com'
        Uri = "$siteUrl/_api/Web/SiteGroups/GetById($groupId)/Users"
        Method = 'GET'
        Headers = $headers
    }

    try {
        $response = Invoke-RestMethod @getUserIdSplatParams           
        return ($response.value | Where-Object { $_.UserPrincipalName -eq $userName}).ID
    }
    catch {
        Write-Error "Failed to retrieve userid of member: $_"
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

    $actionMessage = "add members to group"
    foreach ($user in $membersToAdd) {        
         $addMembersSplatParams = @{
            SharePointToken = $entraToken 
            TenantName = $SharePointBaseUrl        
            UserName = $user.userPrincipalName                       
        }        
        $null = Add-SharePointGroupMembers @addMembersSplatParams

        Write-Information "Successfully added User [$($user.displayName)] to Members of [$($form.sharepointGroups.GroupName)]"
        $Log = @{
            Action            = "GrantMembership" # optional. ENUM (undefined = default) 
            System            = "SharePoint" # optional (free format text) 
            Message           = "Successfully added User [$($user.displayName)] to Members of [$($form.sharepointGroups.GroupName)]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $($user.displayName) # optional (free format text) 
            TargetIdentifier  = $($form.sites.GroupId) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }

    $actionMessage = "remove members from group"
    foreach ($user in $membersToRemove) {
        $getUserIdSplatParams = @{
            SharePointToken = $entraToken 
            TenantName = $SharePointBaseUrl        
            UserName = $user.UserPrincipalName                       
        }

        $userId = Get-SharePointUserIdByUsername @getUserIdSplatParams

        $removeMembersSplatParams = @{
            SharePointToken = $entraToken 
            TenantName = $SharePointBaseUrl        
            UserId = $userid
        }        
        $null = Remove-SharePointGroupMembers @removeMembersSplatParams    
        
        Write-Information "Successfully removed User [$($user.DisplayName)] from Members of [$($form.sharepointGroups.GroupName)]"
        $Log = @{
            Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
            System            = "SharePoint" # optional (free format text) 
            Message           = "Successfully removed User [$($user.DisplayName)] from Members of [$($form.sharepointGroups.GroupName)]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $($user.Name) # optional (free format text) 
            TargetIdentifier  = $($form.sites.GroupId) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log        
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

    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "SharePoint" # optional (free format text) 
        Message           = $auditMessage # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $($form.sharepointGroups.GroupName) # optional (free format text) 
        TargetIdentifier  = $($form.sites.GroupId) # optional (free format text) 
    }
    
    Write-Information -Tags "Audit" -MessageData $log
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
