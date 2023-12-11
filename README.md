# GCX Site Creation

## Summary

Create a new site (community) in GCXchange
- Create an MS Graph group after ensuring the proposed SharePoint url is not already in use
- Update the request with the newly created url
- Create a team and add it to the group
- Associate the site with a hub site
- Apply a provisioning template using PnP Framework
- Add a message to the appropriate sensitivity queue so that appropriate permissions can be applied

If the site already exists then a message is added to the status queue.

## Prerequisites

The following user accounts (as reflected in the app settings) are required:

| Account             | Membership requirements                                  |
| ------------------- | -------------------------------------------------------- |
| delegatedUserName   | Member of site that hosts the list of community requests |
| delegatedUserName   | Member of site that hosts the list of teams links        |

## Version 

![dotnet 6](https://img.shields.io/badge/net6.0-blue.svg)

## API permission

MSGraph

| API / Permissions name    | Type        | Admin consent | Justification                       |
| ------------------------- | ----------- | ------------- | ----------------------------------- |
| Group.Create              | Application | Yes           | Initial group creation              |
| Group.ReadWrite.All       | Delegated   | Yes           | Associated Team creation            | 
| GroupMember.ReadWrite.All | Application | Yes           | Add members to team                 |
| User.Read.All             | Application | Yes           | Retrieve user information           |

Sharepoint

| API / Permissions name    | Type      | Admin consent | Justification                       |
| ------------------------- | --------- | ------------- | ----------------------------------- |
| AllSites.FullControl      | Delegated | Yes           | Apply template                      |

## App setting

| Name                    	| Description                                                                   					         |
| -------------------------	| ------------------------------------------------------------------------------------------------ |
| apprefSiteId              | Id of the SharePoint stei taht hosts the list of teams links |
| AzureWebJobsStorage     	| Connection string for the storage acoount                                     					         |
| clientId                	| The application (client) ID of the app registration                           					         |
| delegatedUserName         | Delegated authentication user for applying the template and updating Request list 				       |
| delegatedUserSecret       | The secret name for the delegated user password 													                       |
| followingContentFeatureId | Id of the Following Content feature used to remove it from the template 							           |
| hubSiteId					        | Id of the hub site that is associated with newly created sites 									                 |
| keyVaultUrl             	| Address for the key vault                                                     					         |
| listId					          | Id of the SharePoint list for community requests                              					         |
| ownerId					          | Id of  the service account to add as temporary owner in order to authorize delegated permissions |
| secretName              	| Secret name used to authorize the function app                                					         |
| sharePointUrl				      | The base url under which new sites will be created 												                       |
| siteId					          | Id of the SharePoint site that hosts the list of community requests           					         |
| teamsLinkListId | Id of the SharePoint list for teams links |
| tenantId                	| Id of the Azure tenant that hosts the function app                            					         |
| tenantName				        | Name of the tenant that hosts the function app                                					         |

## Version history

Version|Date|Comments
-------|----|--------
1.0|2023-10-10|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
