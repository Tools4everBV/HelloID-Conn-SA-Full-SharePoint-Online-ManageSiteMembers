# HelloID-Conn-SA-Full-SharePoint-Online-ManageSiteMembers

| :information_source: Information                                                                                                                                                                                                                                                                                                                                                          |
| :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements. |

| :warning: Important                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Best Practice:** Use **HelloID Products** for requesting and managing permissions (group memberships, mailbox access, application roles). Products provide governance, approval workflows, admin visibility, and full lifecycle management.<br>Use delegated forms for one-time operational actions (creating resources like shared mailboxes, password resets, attribute updates) only.<br><br>**[Read more: Products vs. Delegated Forms](https://docs.helloid.com/en/service-automation/products-vs--delegated-forms.html)** |

| :information_source: Information                                                                                                                                                                                                                                                                                                                                                          |
| :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements. |

## Description

_HelloID-Conn-SA-Full-SharePoint-Online-ManageSiteMembers_ is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements.

This HelloID Service Automation Delegated Form provides SharePoint Online functionality. The following options are available:

1.  Search site based on sitename
2.  Select a site from search results
3.  Select a sitegroup of which the members will be managed
4.  Manage the members of the selected group
5.  Confirm the changes

#### App Registration & Certificate Setup

Before implementing this connector, make sure to configure a Microsoft Entra ID App Registration. During the setup process, you'll create a new App Registration in the Entra portal, assign the necessary API permissions (such as user and group read/write), and generate and assign a certificate.

Follow the official Microsoft documentation for creating an App Registration and setting up certificate-based authentication:

- [App-only authentication with certificate](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication)

#### HelloID-specific configuration

Once you have completed the Microsoft setup and followed their best practices, configure the following HelloID-specific requirements.

- **API Permissions (Application permissions)**:
  - `Group.ReadWrite.All` - To manage group memberships
  - `Sites.ReadWrite.All` - To read site information

- **Certificate Base64 encoded string**:
  - Base64 encoded string of the certificate assigned to the app registration. For instructions on creating the certificate and obtaining the base64 string, refer to our forum post: [Setting up a certificate for Microsoft Graph API in HelloID connectors](https://forum.helloid.com/forum/helloid-provisioning/5338-instruction-setting-up-a-certificate-for-microsoft-graph-api-in-helloid-connectors#post5338)

- **SharePoint REST API**:
  - Besides the use of the Graph API, details of the SharePoint site are retrieved from the REST API of SharePoint itself. To make use of the REST API additional API permissions are needed through the SharePoint API Permissions Scope.
  - `Sites.ReadWrite.All`
  - `Sites.FullControl.All`

- **Application Registration Redirect**:
  - Additional the redirect uri needs to be set on the Application Registratin through the Authentication settings.
    - https://learn.microsoft.com/en-us/entra/identity-platform/how-to-add-redirect-uri
    - Platform: `Mobile and desktop applications`
    - Redirect Uri: `https://localhost`

### Connection settings

The following global variables must be configured in HelloID when importing and configuring the delegated form.

| Setting                        | Description                                                              | Mandatory |
| ------------------------------ | ------------------------------------------------------------------------ | --------- |
| EntraIdTenantId                | The unique identifier (ID) of the tenant in Microsoft Entra ID           | Yes       |
| EntraIdAppId                   | The unique identifier (ID) of the App Registration in Microsoft Entra ID | Yes       |
| EntraIdCertificateBase64String | The Base64-encoded string representation of the app certificate          | Yes       |
| EntraIdCertificatePassword     | The password associated with the app certificate                         | Yes       |
| SharePointBaseUrl              | Sharepoint Base Url for SharePoint REST API                              | Yes       |

## Remarks

- **JWT Token Generation**: The connector uses certificate-based authentication to generate JSON Web Tokens (JWT) for secure communication with Microsoft Graph API. The certificate is converted from a base64 string and used to sign the JWT assertion for OAuth2 authentication.

## Development resources

### API endpoints

The following Microsoft Graph API endpoints are used by the connector:

| Endpoint                   | Description         |
| -------------------------- | ------------------- |
| /v1.0/groups               | List groups         |
| /_SharePointBaseUrl_/\_api | SharePoint REST API |

### API documentation

- [SharePoint API](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service?tabs=csom)
- [List groups](https://learn.microsoft.com/en-us/graph/api/group-list)

## Getting help

| :bulb: Tip                                                                                                                                               |
| :------------------------------------------------------------------------------------------------------------------------------------------------------- |
| For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages. |

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/
