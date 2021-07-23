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
