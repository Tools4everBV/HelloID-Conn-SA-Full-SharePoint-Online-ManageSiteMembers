$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form:
$groupId = $form.sharepointGroups.GroupName
$membersToAdd = $form.members.leftToRight
$membersToRemove = $form.members.rightToLeft
$siteUrl = $form.sites.Url

Write-Verbose "Members to add: $membersToAdd"
Write-Verbose "Members to remove: $membersToRemove"

$connected = $false
try {
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    $pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
    $cred = [System.Management.Automation.PSCredential]::new($SharePointAdminUser, $pwd)
    $null = Connect-SPOService -Url $SharePointBaseUrl -Credential $cred
    Write-Information "Connected to Microsoft SharePoint"
    $connected = $true
}
catch {	
    Write-Error "Could not connect to Microsoft SharePoint. Error: $($_.Exception.Message)"
}

if ($connected) {
    try {
        foreach ($user in $membersToAdd) {
            try {
                $username = $user.User
                $addSPOUser = Add-SPOUser -Site $siteUrl -Group $groupId -LoginName $username
                Write-Information "Successfully added User [$username] to Members of [$groupId]"

                $userDisplayName = $adUser.Name
                $userId = $user.User
                $Log = @{
                    Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                    System            = "SharePoint" # optional (free format text) 
                    Message           = "Successfully added User [$username] to Members of [$groupId]" # required (free format text) 
                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $userDisplayName # optional (free format text) 
                    TargetIdentifier  = $userId # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log
            }
            catch {
                Write-Error "Could not add User [$username] to Members of [$groupId]. Error: $($_.Exception.Message)"

                $userDisplayName = $adUser.Name
                $userId = $user.User
                $Log = @{
                    Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                    System            = "SharePoint" # optional (free format text) 
                    Message           = "Failed to add User [$username] to Members of [$groupId]. Error: $($_.Exception.Message)" # required (free format text) 
                    IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $userDisplayName # optional (free format text) 
                    TargetIdentifier  = $userId # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log
            }
        }

        foreach ($user in $membersToRemove) {
            try {
                $username = $user.User
                $removeSPOUser = Remove-SPOUser -Site $siteUrl -Group $groupId -LoginName $username
                Write-Information "Successfully removed User [$username] from Members of [$groupId]" -Event Success
            
                $userDisplayName = $adUser.Name
                $userId = $user.User
                $Log = @{
                    Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                    System            = "SharePoint" # optional (free format text) 
                    Message           = "Successfully removed User [$username] from Members of [$groupId]" # required (free format text) 
                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $userDisplayName # optional (free format text) 
                    TargetIdentifier  = $userId # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log
            }
            catch {
                Write-Error "Could not remove User [$username] from Members of [$groupId]. Error: $($_.Exception.Message)"
            
                $userDisplayName = $adUser.Name
                $userId = $user.User
                $Log = @{
                    Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                    System            = "SharePoint" # optional (free format text) 
                    Message           = "Failed to remove User [$username] from Members of [$groupId]. Error: $($_.Exception.Message)" # required (free format text) 
                    IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $userDisplayName # optional (free format text) 
                    TargetIdentifier  = $userId # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log
            }
        }   
    }
    finally {
        Disconnect-SPOService
        Remove-Module Microsoft.Online.SharePoint.PowerShell
    }
}
