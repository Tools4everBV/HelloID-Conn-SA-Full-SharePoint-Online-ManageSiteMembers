# Add TLS1.2 support to the script
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityPrototolType]::Tls12
#HelloID variables
$script:PortalBaseUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("Users") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("SharePoint") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> SharePointRootUrl
$tmpName = @'
SharePointRootUrl
'@ 
$tmpValue = @'
https://customer.sharepoint.com
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #2 >> SharePointAdminUser
$tmpName = @'
SharePointAdminUser
'@ 
$tmpValue = @'
user@customer.onmicrosoft.com
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> SharePointAdminPWD
$tmpName = @'
SharePointAdminPWD
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #4 >> SharePointBaseUrl
$tmpName = @'
SharePointBaseUrl
'@ 
$tmpValue = @'
https://customer-admin.sharepoint.com
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});


#make sure write-information logging is visual
$InformationPreference = "continue"
# Create authorization headers with HelloID API key
$pair = "$apiKey" + ":" + "$apiSecret"
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)
$key = "Basic $base64"
$script:headers = @{"authorization" = $Key}
# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"
 
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
            $body = ConvertTo-Json -InputObject $body
    
            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid
            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
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
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}
    
        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task
            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = [Object[]]($Variables | ConvertFrom-Json);
            }
            $body = ConvertTo-Json -InputObject $body
    
            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid
            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
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
        [parameter(Mandatory)][Ref]$returnObject
    )
    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }
    
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
      
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = [Object[]]($DatasourceModel | ConvertFrom-Json);
                automationTaskGUID = $AutomationTaskGuid;
                value              = [Object[]]($DatasourceStaticValue | ConvertFrom-Json);
                script             = $DatasourcePsScript;
                input              = [Object[]]($DatasourceInput | ConvertFrom-Json);
            }
            $body = ConvertTo-Json -InputObject $body
      
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
              
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
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
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = [Object[]]($FormSchema | ConvertFrom-Json)
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }
    $returnObject.Value = $formGuid
}
function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][String][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                accessGroups    = [Object[]]($AccessGroups | ConvertFrom-Json);
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
            }    
            $body = ConvertTo-Json -InputObject $body
    
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true
            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }
    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "Sharepoint-get-site-group-members" #>
$tmpPsScript = @'
$siteUrl = $datasource.selectedSite.Url
$siteName = $datasource.selectedSite.DisplayName
$group = $datasource.selectedGroup.GroupName
$connected = $false

try {
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	$pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $SharePointAdminUser, $pwd
	$null = Connect-SPOService -Url $SharePointBaseUrl -Credential $cred
    Write-Information "Connected to Microsoft SharePoint"
    $connected = $true
}
catch
{	
    Write-Error "Could not connect to Microsoft SharePoint. Error: $($_.Exception.Message)"
    Write-Warning "Failed to connect to Microsoft SharePoint"
}

if ($connected)
{
    try {
        Write-Information "$siteUrl"
        $sharePointUsers = Get-SPOUser -Site $siteUrl -Group "$group"
        Write-Information $sharePointUsers.Count
        foreach($spUser in $sharePointUsers)
            {
                $returnObject = @{Name=$spUser.DisplayName; User=$spUser.LoginName}
                Write-Output $returnObject
            }
        
	}
	catch
	{
		Write-Error "Error getting Site Details. Error: $($_.Exception.Message)"
		Write-Warning -Message "Error getting Site Details"
		return
	}
    finally
    {        
        Disconnect-SPOService
        Remove-Module -Name Microsoft.Online.SharePoint.PowerShell
    }
}
else
{
	return
}

'@ 
$tmpModel = @'
[{"key":"User","type":0},{"key":"Name","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedSite","type":0,"options":1},{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedGroup","type":0,"options":1}]
'@ 
$dataSourceGuid_3 = [PSCustomObject]@{} 
$dataSourceGuid_3_Name = @'
Sharepoint-get-site-group-members
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_3_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_3) 
<# End: DataSource "Sharepoint-get-site-group-members" #>

<# Begin: DataSource "Sharepoint-get-site-groups" #>
$tmpPsScript = @'
$siteId = $datasource.selectedSite.Url
$connected = $false

try {
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	$pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $SharePointAdminUser, $pwd
	$null = Connect-SPOService -Url $SharePointBaseUrl -Credential $cred
    Write-Information "Connected to Microsoft SharePoint"
    $connected = $true
}
catch
{	
    Write-Error "Could not connect to Microsoft SharePoint. Error: $($_.Exception.Message)"
    Write-Warning "Failed to connect to Microsoft SharePoint"
}

if ($connected)
{
	try {
        $groups = Get-SPOSiteGroup -Site $siteId
        $GroupsData = @()
        Write-Information $sites.Count
        if(@($sites).Count -ge 0){
         foreach($Group in $groups)
            {                   
                $returnObject = [ordered]@{GroupName=$Group.Title}
                Write-Output $returnObject                
            }
        }else{
            return
        }
	}
	catch
	{
		Write-Error "Error getting Site Details. Error: $($_.Exception.Message)"
		Write-Warning -Message "Error getting Site Details"
		return
	}
    finally
    {
        Disconnect-SPOService
        Remove-Module -Name Microsoft.Online.SharePoint.PowerShell
    }
}
else
{
	return
}

'@ 
$tmpModel = @'
[{"key":"GroupName","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedSite","type":0,"options":1}]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
Sharepoint-get-site-groups
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "Sharepoint-get-site-groups" #>

<# Begin: DataSource "Sharepoint-get-azure-users" #>
$tmpPsScript = @'
# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try {
        Write-Information "Generating Microsoft Graph API Access Token user.."

        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$AADTenantID/oauth2/token"
        
        $body = @{
            grant_type      = "client_credentials"
            client_id       = "$AADAppId"
            client_secret   = "$AADAppSecret"
            resource        = "https://graph.microsoft.com"
        }
 
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
        $accessToken = $Response.access_token;
        Write-Information $accessToken
        #Add the authorization header to the request
        $authorization = @{
            Authorization = "Bearer $accesstoken";
            'Content-Type' = "application/json";
            Accept = "application/json";
        }

        $propertiesToSelect = @(
            "UserPrincipalName",
            "GivenName",
            "Surname",
            "EmployeeId",
            "AccountEnabled",
            "DisplayName",
            "OfficeLocation",
            "Department",
            "JobTitle",
            "Mail",
            "MailNickName"            
        )
 
        $usersUri = "https://graph.microsoft.com/v1.0/users"        
        $usersUri = $usersUri + ('?$select=' + ($propertiesToSelect -join "," | Out-String))
        
        $data = @()
        $query = Invoke-RestMethod -Method Get -Uri $usersUri -Headers $authorization -ContentType 'application/x-www-form-urlencoded'
        $data += $query.value
        
        while($query.'@odata.nextLink' -ne $null){
            $query = Invoke-RestMethod -Method Get -Uri $query.'@odata.nextLink' -Headers $authorization -ContentType 'application/x-www-form-urlencoded'
            $data += $query.value 
        }
        
        $users = $data #| Sort-Object -Property DisplayName
        $resultCount = @($users).Count
        Write-Information "Result count: $resultCount"        
          
        if($resultCount -gt 0){
            foreach($user in $users){  
                if($user.UserPrincipalName -ne $SharePointAdminUser)    
                {          
                    $returnObject = @{User=$user.UserPrincipalName; Name=$user.displayName}
                    Write-Output $returnObject                
                }
            }
        } else {
            return
        }
    }
 catch {
    $errorDetailsMessage = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message
    Write-Error ("Error searching for AzureAD users. Error: $($_.Exception.Message)" + $errorDetailsMessage)
     
    return
}
'@ 
$tmpModel = @'
[{"key":"User","type":0},{"key":"Name","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_2 = [PSCustomObject]@{} 
$dataSourceGuid_2_Name = @'
Sharepoint-get-azure-users
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_2_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_2) 
<# End: DataSource "Sharepoint-get-azure-users" #>

<# Begin: DataSource "Sharepoint-generate-table-sites-wildcard" #>
$tmpPsScript = @'
$connected = $false
$searchValue = $datasource.searchValue
try {
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	$pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $SharePointAdminUser, $pwd
	$null = Connect-SPOService -Url $SharePointBaseUrl -Credential $cred
    Write-Information "Connected to Microsoft SharePoint"
    $connected = $true
}
catch
{	
    Write-Error "Could not connect to Microsoft SharePoint. Error: $($_.Exception.Message)"
    Write-Warning "Failed to connect to Microsoft SharePoint"
}

if ($connected)
{    
	try {
        #Write-Output $searchValue
	    $sites = Get-SPOSite -Filter "url -like 'sites/$($searchValue)'" -Limit ALL

       ForEach($Site in $sites)
        {
            #Write-Output $Site 
            $returnObject = @{DisplayName=$Site.Title; Url=$Site.Url;}
            Write-Output $returnObject                
        }
        
	}
	catch
	{
		Write-Error "Error getting SharePoint sitecollections. Error: $($_.Exception.Message)"
		Write-Warning "Error getting SharePoint sitecollections"
		return
	}
    finally
    {        
        Disconnect-SPOService
        Remove-Module -Name Microsoft.Online.SharePoint.PowerShell
    }
}
else
{
	return
}

'@ 
$tmpModel = @'
[{"key":"DisplayName","type":0},{"key":"Url","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"searchValue","type":0,"options":1}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
Sharepoint-generate-table-sites-wildcard
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "Sharepoint-generate-table-sites-wildcard" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "SharePoint - Manage Site Members" #>
$tmpSchema = @"
[{"label":"Select Site","fields":[{"key":"searchValue","templateOptions":{"label":"Search for SiteName","required":true},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"sites","templateOptions":{"label":"Select Site","required":true,"grid":{"columns":[{"headerName":"Display Name","field":"DisplayName"},{"headerName":"Url","field":"Url"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[{"propertyName":"searchValue","otherFieldValue":{"otherFieldKey":"searchValue"}}]}},"useDefault":false,"useFilter":true},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]},{"label":"Site groups","fields":[{"key":"sharepointGroups","templateOptions":{"label":"Select Group","required":true,"grid":{"columns":[{"headerName":"Group Name","field":"GroupName"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[{"propertyName":"selectedSite","otherFieldValue":{"otherFieldKey":"sites"}}]}},"useDefault":false},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]},{"label":"Manage Members","fields":[{"key":"members","templateOptions":{"label":"Manage members","required":false,"filterable":false,"useDataSource":true,"dualList":{"options":[{"guid":"75ea2890-88f8-4851-b202-626123054e14","Name":"Apple"},{"guid":"0607270d-83e2-4574-9894-0b70011b663f","Name":"Pear"},{"guid":"1ef6fe01-3095-4614-a6db-7c8cd416ae3b","Name":"Orange"}],"optionKeyProperty":"Name","optionDisplayProperty":"Name","labelLeft":"Available","labelRight":"Current"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_2","input":{"propertyInputs":[]}},"destinationDataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_3","input":{"propertyInputs":[{"propertyName":"selectedSite","otherFieldValue":{"otherFieldKey":"sites"}},{"propertyName":"selectedGroup","otherFieldValue":{"otherFieldKey":"sharepointGroups"}}]}}},"type":"duallist","summaryVisibility":"Show","sourceDataSourceIdentifierSuffix":"source-datasource","destinationDataSourceIdentifierSuffix":"destination-datasource","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
SharePoint - Manage Site Members
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
foreach($group in $delegatedFormAccessGroupNames) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $delegatedFormAccessGroupGuid = $response.groupGuid
        $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
        
        Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
    } catch {
        Write-Error "HelloID (access)group '$group', message: $_"
    }
}
$delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Compress)
$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
SharePoint - Manage Site Members
'@
Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-puzzle-piece" -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

<# Begin: Delegated Form Task #>
if($delegatedFormRef.created -eq $true) { 
	$tmpScript = @'
HID-Write-Status -Message "Members to add: $MembersToAdd" -Event Information
HID-Write-Status -Message "Members to remove: $MembersToRemove" -Event Information

$connected = $false
try {
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	$pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $SharePointAdminUser, $pwd
	$null = Connect-SPOService -Url $SharePointBaseUrl -Credential $cred
    HID-Write-Status -Message "Connected to Microsoft SharePoint" -Event Information
    HID-Write-Summary -Message "Connected to Microsoft SharePoint" -Event Information
	$connected = $true
}
catch
{	
    HID-Write-Status -Message "Could not connect to Microsoft SharePoint. Error: $($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Failed to connect to Microsoft SharePoint" -Event Failed
}

if ($connected)
{
	try{
        if($MembersToAdd -ne "[]"){
            HID-Write-Status -Message "Starting to add Users to Members of [$groupId]: $MembersToAdd" -Event Information
            $usersToAddJson =  $MembersToAdd | ConvertFrom-Json
            
            foreach($user in $usersToAddJson)
            {
                try{
                    $username = $user.User
                    Add-SPOUser -Site $siteUrl -Group $groupId -LoginName $username
                    HID-Write-Status -Message "Finished adding User [$username] to Members of [$groupId]" -Event Success
                    HID-Write-Summary -Message "Successfully added User [$username] to Members of [$groupId]" -Event Success
                }
                catch{
                    HID-Write-Status -Message "Could not add User [$username] to Members of [$groupId]. Error: $($_.Exception.Message)" -Event Error
                    HID-Write-Summary -Message "Failed to add User [$username] to Members of [$groupId]" -Event Failed
                }
            }
        }
        
        if($MembersToRemove -ne "[]"){
            HID-Write-Status -Message "Starting to remove Users to Members of [$groupId]: $MembersToRemove" -Event Information
            $usersToRemoveJson =  $MembersToRemove | ConvertFrom-Json
                
            foreach($user in $usersToRemoveJson)
            {
                try{
                    $username = $user.User
                    Remove-SPOUser -Site $siteUrl -Group $groupId -LoginName $username
                    HID-Write-Status -Message "Finished removing User [$username] from Members of [$groupId]" -Event Success
                    HID-Write-Summary -Message "Successfully removed User [$username] from Members of [$groupId]" -Event Success
                }
                catch{
                    HID-Write-Status -Message "Could not remove User [$username] from Members of [$groupId]. Error: $($_.Exception.Message)" -Event Error
                    HID-Write-Summary -Message "Failed to remove User [$username] from Members of [$groupId]" -Event Failed
                }
            }   
        }
    }
    catch
	{
		HID-Write-Status -Message "Could not manage members of group [$groupId]. Error: $($_.Exception.Message)" -Event Error
		HID-Write-Summary -Message "Failed to manage members of group [$groupId]" -Event Failed
	}
    finally
    {
        Disconnect-SPOService
        Remove-Module Microsoft.Online.SharePoint.PowerShell
    }
}
'@; 

	$tmpVariables = @'
[{"name":"groupId","value":"{{form.sharepointGroups.GroupName}}","secret":false,"typeConstraint":"string"},{"name":"MembersToAdd","value":"{{form.members.leftToRight.toJsonString}}","secret":false,"typeConstraint":"string"},{"name":"MembersToRemove","value":"{{form.members.rightToLeft.toJsonString}}","secret":false,"typeConstraint":"string"},{"name":"siteUrl","value":"{{form.sites.Url}}","secret":false,"typeConstraint":"string"}]
'@ 

	$delegatedFormTaskGuid = [PSCustomObject]@{} 
$delegatedFormTaskName = @'
SharePoint - Manage SharePoint Members
'@
	Invoke-HelloIDAutomationTask -TaskName $delegatedFormTaskName -UseTemplate "False" -AutomationContainer "8" -Variables $tmpVariables -PowershellScript $tmpScript -ObjectGuid $delegatedFormRef.guid -ForceCreateTask $true -returnObject ([Ref]$delegatedFormTaskGuid) 
} else {
	Write-Warning "Delegated form '$delegatedFormName' already exists. Nothing to do with the Delegated Form task..." 
}
<# End: Delegated Form Task #>
