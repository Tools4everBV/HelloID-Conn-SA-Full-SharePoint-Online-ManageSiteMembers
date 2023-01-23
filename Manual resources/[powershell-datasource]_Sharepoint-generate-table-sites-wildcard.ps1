$connected = $false
$searchValue = $datasource.searchValue
try {
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	$pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
	$cred = [System.Management.Automation.PSCredential]::new($SharePointAdminUser,$pwd)
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

