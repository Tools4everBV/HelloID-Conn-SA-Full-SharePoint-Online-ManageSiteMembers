$siteId = $datasource.selectedSite.Url
$connected = $false

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

