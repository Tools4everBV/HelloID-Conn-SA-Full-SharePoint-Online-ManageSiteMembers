<!-- Description -->
## Description
This HelloID Service Automation Delegated Form provides SharePoint Online functionality. The following options are available:
 1. Search site based on sitename
 2. Select a site from search results
 3. Select a sitegroup of which the members will be managed
 4. Manage the members of the selected group
 5. Confirm the changes
 
<!-- TABLE OF CONTENTS -->
## Table of Contents
* [Description](#description)
* [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  * [Getting started](#getting-started)
* [Prerequisites](#prerequisites)
* [Post-setup configuration](#post-setup-configuration)
* [Manual resources](#manual-resources)


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

 _Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_


### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.

## Prerequisites
Please follow the documentation for installing the SharePoint powershell module on [Microsoft Docs](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps).
 
## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)
<table>
  <tr><td><strong>Variable name</strong></td><td><strong>Example value</strong></td><td><strong>Description</strong></td></tr>
  <tr><td>SharePointRootUrl</td><td>https://customer.sharepoint.com</td><td>URL of customers SharePoint environment</td></tr>
  <tr><td>SharePointBaseUrl</td><td>https://customer-admin.sharepoint.com</td><td>URL of customers admin SharePoint environment</td></tr>
  <tr><td>SharePointAdminUser</td><td>svc_account@customer.onmicrosoft.com</td><td>Azure user account with the SharePoint Admin Role</td></tr>
  <tr><td>SharePointAdminPWD</td><td>********</td><td>SharePointAdminUser password</td></tr>
</table>

## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source '[powershell-datasource]_Sharepoint-generate-table-sites-wildcard'
This Powershell data source performs a search on available SharePoint sites.

### Powershell data source '[powershell-datasource]_Sharepoint-get-azure-users'
This Powershell data source gets all available Azure users to select a site owner.

### Powershell data source '[powershell-datasource]_Sharepoint-get-site-groups'
This Powershell data source performs a search on available groups within the selected SharePoint site.

### Powershell data source '[powershell-datasource]_Sharepoint-get-site-group-members'
This Powershell data source gets members of the selected SharePoint group.

### Delegated form task '[task]_SharePoint - Manage SharePoint Members'
This delegated form task will add and/or remove members from the selected SharePoint group.

# HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/
