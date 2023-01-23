$siteUrl = $datasource.selectedSite.Url
$siteName = $datasource.selectedSite.DisplayName
$group = $datasource.selectedGroup.GroupName
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

