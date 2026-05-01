# IA Governance Extension - Function Library

**Last Updated:** 01-05-2026 | **Version:** 2.0 | **Scope:** Generic Functions Only

> This documentation provides comprehensive information about all available functions and actions in the IA Governance Extension library.

## 📑 Table of Contents

### Configuration Functions

- [ApplyAutoFillColumnsSettings](#applyautofillcolumnssettings)
- [ApplyFormCustomizer](#applyformcustomizer)
- [ApplySiteMetadata-deprecated](#applysitemetadata-deprecated)
- [ApplyTheme](#applytheme)
- [SetColumnDefaultValue_Multi](#setcolumndefaultvalue_multi)
- [SetColumnDefaultValue-deprecated](#setcolumndefaultvalue-deprecated)
- [SetColumnDefaultValueTaxonomy_nonIAG](#setcolumndefaultvaluetaxonomy_noniag)
- [SetColumnDefaultValueTaxonomy_nonIAG_Multi](#setcolumndefaultvaluetaxonomy_noniag_multi)
- [SetColumnDefaultValueTaxonomy-deprecated](#setcolumndefaultvaluetaxonomy-deprecated)
- [SetDefaultValueOnContentType](#setdefaultvalueoncontenttype)
- [SetDocumentIDPrefix](#setdocumentidprefix)
- [SetExternalUserExpirationInDays](#setexternaluserexpirationindays)
- [SetHeaderTitleVisibility](#setheadertitlevisibility)
- [SetHTMLFieldSecurity](#sethtmlfieldsecurity)
- [SetHubSiteAssociation](#sethubsiteassociation)
- [SetLibraryVersionLimit](#setlibraryversionlimit)
- [SetLogo](#setlogo)
- [SetM365EmailProperties](#setm365emailproperties)
- [SetM365SMTPEmail](#setm365smtpemail)
- [SetMembersToContribute](#setmemberstocontribute)
- [SetNavigation](#setnavigation)
- [SetPermissionForSPGroup](#setpermissionforspgroup)
- [SetRegionalSettings](#setregionalsettings)
- [SetSiteAdmins](#setsiteadmins)
- [SetSiteMembers](#setsitemembers)
- [SetSiteOwners](#setsiteowners)
- [SetSiteVisitors](#setsitevisitors)
- [SetSPAccessRequest](#setspaccessrequest)
- [SetSPFieldProperties](#setspfieldproperties)
- [SetSPGroupNames](#setspgroupnames)
- [SetStorageLimit](#setstoragelimit)
- [SetTeamsChannelSiteOwners](#setteamschannelsiteowners)
- [UpdateCalculatedFieldFormula](#updatecalculatedfieldformula)
- [UpdateIAGExtensionLogItem](#updateiagextensionlogitem)
- [UpdateListColumnDescription](#updatelistcolumndescription)
- [UpdateMultiLineFieldToUseEnhancedRichText](#updatemultilinefieldtouseenhancedrichtext)
- [UpdatePortalItem](#updateportalitem)

### Creation Functions

- [AddColumnFormatting](#addcolumnformatting)
- [AddCustomActionToGearmenu](#addcustomactiontogearmenu)
- [AddCustomPermissionLevel](#addcustompermissionlevel)
- [AddFoldersToLibrary](#addfolderstolibrary)
- [AddFoldersToLibraryInBatch -deprecated](#addfolderstolibraryinbatch--deprecated)
- [AddFunctionToGearmenu -deprecated](#addfunctiontogearmenu--deprecated)
- [AddGroupOwnersAsChannelOwners](#addgroupownersaschannelowners)
- [AddIAGExtensionLogItem](#addiagextensionlogitem)
- [AddM365GroupOwner](#addm365groupowner)
- [AddNewFieldFromXML](#addnewfieldfromxml)
- [AddPlannerPlan](#addplannerplan)
- [AddSPGroupsWithPermissions](#addspgroupswithpermissions)
- [AddTeamsChannel](#addteamschannel)
- [AddTerm](#addterm)
- [CreateM365TeamsTab](#createm365teamstab)
- [CreateM365TeamsTab_Custom](#createm365teamstab_custom)
- [CreateM365TeamsTab_DocumentLibrary](#createm365teamstab_documentlibrary)
- [CreateM365TeamsTab_OneNote](#createm365teamstab_onenote)
- [CreateM365TeamsTab_Planner](#createm365teamstab_planner)
- [CreateM365TeamsTab_SharePointPageAndList](#createm365teamstab_sharepointpageandlist)
- [CreateM365TeamsTab_WebSite](#createm365teamstab_website)

### Enablement Functions

- [EnableDisableGridEditingOnList](#enabledisablegrideditingonlist)

### General Functions

- [ActivateFeature](#activatefeature)
- [ActivateWebFeature](#activatewebfeature)
- [ArchiveSiteUsingMSArchive](#archivesiteusingmsarchive)
- [AreContentTypesFromHubReady](#arecontenttypesfromhubready)
- [EviDocCopyFolders](#evidoccopyfolders)
- [EviDocSetColumnDefaultValues](#evidocsetcolumndefaultvalues)
- [EviDokCopyStandardDocsIntoFolders](#evidokcopystandarddocsintofolders)
- [HideListOrLibrary](#hidelistorlibrary)
- [InstallApp](#installapp)
- [InstallAppToTeam](#installapptoteam)
- [NewMicrosoft365Group](#newmicrosoft365group)
- [NewTeamsTeam](#newteamsteam)
- [NewTenantSite](#newtenantsite)
- [ReadFromCreateQueueAndStartFlow](#readfromcreatequeueandstartflow)
- [SendEmailSharePoint-deprecated](#sendemailsharepoint-deprecated)
- [TeamifyMicrosoft365Group](#teamifymicrosoft365group)
- [aa_TestConnection](#aa_testconnection)

### Query Functions

- [GetContentTypesFromHub](#getcontenttypesfromhub)
- [GetIAGMetadata](#getiagmetadata)
- [GetPnPConnection](#getpnpconnection)
- [GetSiteCollectionUrlFromGroupId](#getsitecollectionurlfromgroupid)
- [GetTeamsChannel](#getteamschannel)
- [GetWebIdFromGroupId](#getwebidfromgroupid)

### Removal Functions

- [DeleteContentTypeFromLibrary-deprecated](#deletecontenttypefromlibrary-deprecated)
- [DeleteContentTypesFromList](#deletecontenttypesfromlist)
- [DeleteList](#deletelist)
- [DeleteViewsFromLists](#deleteviewsfromlists)
- [DisableListComments](#disablelistcomments)
- [DisableNextStepsFirstRunEnabled](#disablenextstepsfirstrunenabled)
- [DisableSocialBarOnSitePages](#disablesocialbaronsitepages)
- [RemoveM365GroupOwner](#removem365groupowner)
- [RemovePageBanner](#removepagebanner)

### Templates

- [AddDocumentTemplateToLibrary](#adddocumenttemplatetolibrary)
- [ApplyPnPTemplate](#applypnptemplate)
- [ApplyPnPTemplateFromXML](#applypnptemplatefromxml)
- [CopyCustomPageTemplates](#copycustompagetemplates)
- [DisableWebTemplatesGalleryFirstRunDialog](#disablewebtemplatesgalleryfirstrundialog)



---

## SetDocumentIDPrefix

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

requies that the site collection feature "Document ID Service" is activated

### 💡 Examples

**Example 1**
```powershell
{
  "DocIDPrefix": "A-Collab",
  "siteUrl": "https://contoso.sharepoint.com/sites/contoso" 
}
```

---

## SetDefaultValueOnContentType

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
    "siteUrl": "https://banedanmarkonline.sharepoint.com/sites/ANL10709"
    "ContentTypeName": "NGContractDocumentSet",
    "ListName" : "Contracts",
    "InternalColumnName": "NGProjectID",
    "defaultValue": "test value"
}
```

---

## SetColumnDefaultValue_Multi

**Version:** 1.0 | 
### 📖 Description

Based on the intern field name the type of field is determined, and the value is set accordingly.
    The function can handle both single-value and multi-value taxonomy fields.

### 💡 Examples

**Example 1**
```powershell
{
    "LibraryColumnValues": [
            {
            "LibraryColumnValue": [
                {
                    "Library" : "Test2",
                    "Column" : "MultiValueTaxonomyField",
                    "FolderName" : "Folder1",
                    "Value" : "81273285-2566-4cc4-aa14-2ec126741201;1b00f506-219c-4188-b848-041477fdb23a"
                },
                {
                    "Column": "bdkOrganization",
                    "Library": "Test2",
                    "FolderName" : "Folder1",
                    "Value": "e5a2b490-eedc-409f-a011-8cb256d67901"
                },
                {
                    "Column": "bdkSiteID",
                    "Library": "Test2",
                    "FolderName" : "Folder1",
                    "Value": "WorkspacePermalink"
                },
                {
                    "Column": "bdkDocumentType",
                    "Library": "BDK Internt",
                    "Value": "#Kontraktopfølgning|60905ca5-e8d1-47ce-bca9-aea3b9c55927"
                },
                {
                    "Column": "bdkProjectType",
                    "Library": "BDK Internt",
                    "Value": "#Infrastrukturprojekt|3503b396-e716-49aa-96df-d3aecee5d986"
                }
            ]
        }
    ],
     "siteUrl": "https://banedanmarkonline.sharepoint.com/sites/ANL10709"
}
```

---

## SetColumnDefaultValueTaxonomy_nonIAG_Multi

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
    "LibraryColumnValues": [
        {
            "LibraryColumnValue": [
                {
                    "Column": "bdkOrganization",
                    "Library": "Notater",
                    "Value": "bdkOrganization",
                    "FolderName":"00. Sandboxes"
                },
                {
                    "Column": "bdkSiteID",
                    "Library": "Notater",
                    "Value": "WorkspacePermalink",
                    "FolderName":"00. Sandboxes"
                },
                {
                    "Column": "bdkDocumentType",
                    "Library": "BDK Internt",
                    "Value": "#Kontraktopfølgning|60905ca5-e8d1-47ce-bca9-aea3b9c55927",
                    "FolderName":"00. Sandboxes"
                },
                {
                    "Column": "bdkProjectType",
                    "Library": "BDK Internt",
                    "Value": "#Infrastrukturprojekt|3503b396-e716-49aa-96df-d3aecee5d986",
                    "FolderName":"00. Sandboxes"
                }
            ]
        }
    ],
    "siteUrl": "https://banedanmarkonline.sharepoint.com/sites/ANL10709"
}
```

---

## SetColumnDefaultValueTaxonomy_nonIAG


### 📖 Description

This function sets the default column value on a list column from a taxonomy term

### 💡 Examples

**Example 1**
```powershell
{
  "ListName": "Shared Documents",
  "SiteColumnInternalName":"Department",
  "FolderName":"00. Sandboxes",
  "siteUrl":"the site url" ,
    "defaultValue":"Group|Set|Finance"  
}
```

---

## SetColumnDefaultValueTaxonomy-deprecated

⚠️ **DEPRECATED**


### 📖 Description

⚠️ **DEPRECATED** - This function is deprecated and should not be used in new implementations.

Precondition: the GetIAGMetadataTags function must be called first to get the metadata tags from IAG.

    This function sets the default column value on a list column from a taxonomy term

    NOTE: This function works for both SharePoint sites and Microsoft 365 groups.If you supply a siteUrl then the function knows that this is not a group based site.

### 💡 Examples

**Example 1**
```powershell
{
  
  "siteName":"test", #optional
  "ListName": "Shared Documents",
  "SiteColumnInternalName":"Department",
  "FolderName":"00. Sandboxes",
  "IAGtargetFieldName":"department",
  "siteUrl":"the site url" ,
  "IAGMetadata":@{body('IAGovernanceFunctionApp-GetIAGMetadata')}
}
  or

{

  "siteName":"test", #optional
  "ListName": "Shared Documents",
  "SiteColumnInternalName":"Department",
  "FolderName":"00. Sandboxes",
  "IAGtargetFieldName":"department",
  "siteUrl":"https://tcwlv.sharepoint.com/sites/bdkorganisation-15005",
  "IAGMetadata":@{body('IAGovernanceFunctionApp-GetIAGMetadata')}
}
```

---

## SetColumnDefaultValue-deprecated

⚠️ **DEPRECATED**

**Version:** 1.0 | 
### 📖 Description

⚠️ **DEPRECATED** - This function is deprecated and should not be used in new implementations.

A longer description of the function, its purpose, common use cases, etc.

### 💡 Examples

**Example 1**
```powershell
example 1
{
  "Field": "search_TextFieldSingle",
  "FieldValue": "test",
  "ListName": "Shared Documents",  
  "siteUrl": "URL"
    }
example 2
    {
  "Field": "search_TextFieldSingle",
  "FieldValue": "test",
  "ListName": "Shared Documents",  
  "siteUrl": "URL",
  "FolderName": "FolderName"
    }
```

---

## SendEmailSharePoint-deprecated

⚠️ **DEPRECATED**

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

⚠️ **DEPRECATED** - This function is deprecated and should not be used in new implementations.



### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "sender":"noreply@email.com",
  "receivers":"joe@email.com;Jannet@email.com",
  "subject":"Test message",
  "emailbody":"This is a test message"
 "emailbodyHMTLpath": "/sites/SampleTeamSite/Emailtemplates/emailtemplate.html", #optional
 "HTMLFileSiteUrl": "https://tcwlv.sharepoint.com/sites/SampleTeamSite/" #optional
}
```

---

## RemovePageBanner

**Version:** 1.0 | 
### 📖 Description

Removes the banner from a SharePoint page by changing its layout type to "Home". This is 
    accomplished by updating the PageLayoutType field in the specified list (typically SitePages) 
    and publishing the page with the change.
    
    HttpStatusCode 200 : Page banner successfully removed
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "ListName": "SitePages",
  "PageName": "Welcome.aspx",
  "siteUrl": "siteUrl"
}
```

---

## SetExternalUserExpirationInDays


### 📖 Description

Configures the number of days before external users are automatically removed from a SharePoint 
    site. This security feature helps maintain control over external user access by requiring 
    periodic review and renewal of external sharing.
    
    HttpStatusCode 200 : Expiration period successfully set
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://fmdkbdkdev01.sharepoint.com/sites/2406A",
    "numberOfDays": 30
}
```

---

## RemoveM365GroupOwner

✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "OwnerUPN": "admin@contoso.com",  #only required for private and shared channels
}
```

---

## NewTenantSite

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "sitetemplate": "STS#0",
  "title": "Fellowmind test site",
  "SPOAdminUrl": "https://tcwlv-admin.sharepoint.com",
  "language" : 1033,
  "timeZone" : 3,
  "siteDesign" : "Topic",
  "siteCollectionAdmin" : "ProActiveAdmin@AmpelmannOperations.onmicrosoft.com",
  "requestedsiteUrl" : "https://tcwlv.sharepoint.com/sites/FellowmindTestSite"
}
```

---

## NewTeamsTeam

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Creates a new Teams team with specified properties. Optionally adds owners and members.

### 💡 Examples

**Example 1**
```powershell
{
  "displayName": "Fellowmind Test Team",
  "mailNickName": "FellowmindTestTeam",
  "description": "A test team for Fellowmind",
  "visibility": "Private",
  "classification": "Confidential",
  "allowAddRemoveApps": true,
  "allowChannelMentions": true,
  "allowCreateUpdateChannels": true,
  "allowCreateUpdateRemoveConnectors": true,
  "allowCreateUpdateRemoveTabs": true,
  "allowCustomMemes": true,
  "allowDeleteChannels": true,
  "allowGiphy": true,
  "allowGuestCreateUpdateChannels": true,
  "allowGuestDeleteChannels": true,
  "allowOwnerDeleteMessages": true,
  "allowStickersAndMemes": true,
  "allowTeamMentions": true,
  "allowUserDeleteMessages": true,
  "allowUserEditMessages": true,
  "giphyContentRating": "Moderate",
  "showInTeamsSearchAndSuggestions": true,
  "owners": ["admin@contoso.com"],
  "members": ["user1@contoso.com", "user2@contoso.com"],
  ,
  "resourceBehaviorOptions": [
    "AllowOnlyMembersToPost",
    "HideGroupInOutlook",
    "WelcomeEmailDisabled"
  ]
  "sensitivityLabels": "d290f1ee-6c54-4b01-90e6-d701748f0851"
}
```

---

## NewMicrosoft365Group

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Creates a new Microsoft 365 Group with specified properties. Optionally adds owners and members.

### 💡 Examples

**Example 1**
```powershell
{
  "displayName": "Finance Department",
  "mailNickName": "FinanceDept",
  "description": "Microsoft 365 Group for Finance team",
  "isPrivate": true,
  "classification": "Confidential",
  "owners": ["manager@contoso.com"],
  "members": ["user1@contoso.com", "user2@contoso.com"],
  ,
  "resourceBehaviorOptions": [
    "AllowOnlyMembersToPost",
    "HideGroupInOutlook",
    "WelcomeEmailDisabled"
  ]
}
```

---

## InstallAppToTeam

**Version:** 1.0 | 
### 📖 Description

Installs a new app to a team
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
1
{
"groupId" :"add6c948-e3e7-4570-8d08-fcf44845fe38",
"AppID" :"",
```

---

## InstallApp

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

InstallApp is a function that will install one or more apps to a site collection.
Each app is identified by its App ID (GUID) and separated by a semicolon.
The apps must already be available in the tenant app catalog.

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://tcwlv.sharepoint.com/sites/bdkorganisation-15007",
  "AppIds": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx;yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
}
```

---

## HideListOrLibrary

**Version:** 1.0 | 
### 📖 Description

Hides a named list from a site.
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  
  "ListName" : "Documents",
  "siteUrl": "the site url" 
}
```

---

## GetWebIdFromGroupId

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Looks up the site collection url from the group id

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "theGUID"
}
```

---

## GetTeamsChannel

✅ Supports Managed Identity and App Only connections

### 📖 Description

AddTeamsChannel is a function that will add a channel to a team.
HttpStatusCode 200 : the channel has been created
HttpStatusCode 208 : the channels exists allready

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "siteName": "@{body('Parse_JSON')?['siteName']}",
  "DisplayName": "test" 
}
```

---

## GetSiteCollectionUrlFromGroupId

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Looks up the site collection url from the group id

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "theGUID"
}
```

---

## ReadFromCreateQueueAndStartFlow

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

This timer-triggered function runs on a configurable schedule configured via the QueueTriggerSchedule
    environment variable, reads messages from a storage queue (configured via QueueName env variable),
    and calls a Logic App workflow for each message by making an HTTP POST request.
    Uses user-assigned managed identity to authenticate with Azure Storage Queue.

---

## SetHeaderTitleVisibility

**Version:** 1.0 | 
### 📖 Description

This is a workaround as this should be in the site provisioning template but the template is of an older version

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://contoso.sharepoint.com/sites/contoso"
}
```

---

## SetHTMLFieldSecurity

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Updates the site's HTML Field Security safe domain list by adding the provided domain
    to CustomScriptSafeDomains. If the domain already exists, no change is made.

### 💡 Examples

**Example 1**
```powershell
{
    "siteUrl": "https://contoso.sharepoint.com/sites/contoso",
    "domainNames": [
        "products.contoso.com",
        "portal.contoso.com"
    ]
}
```

---

## SetHubSiteAssociation

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Adds the current site collection to the Hub site specified in the body (hubSiteUrl)
 If the siteUrl is identical to the hubSiteUrl, then the site is registered as a hub site

### 💡 Examples

**Example 1**
```powershell
{
  
  "groupId": "5012da92-7d94-4cb7-8aa8-26e917fe2562",
  "hubSiteUrl": "https://tcwlv.sharepoint.com/sites/HubsiteA"
 
}
OR
{   
    "hubSiteUrl": "https://tcwlv.sharepoint.com/sites/HubsiteA",
    "siteUrl": "https://contoso.sharepoint.com/sites/contoso"
}
```

---

## UpdateListColumnDescription


### 📖 Description

Updates the description of a column in a SharePoint list. This is useful when you cannot 
    update the description of a site column directly, and need to update the list column instead. 
    This is often the case with system columns like Title.
    
    HttpStatusCode 200 : Column description successfully updated
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
   "siteUrl": "https://contoso.sharepoint.com/sites/contoso",   
  "ListName" : "Documents",
  "ListColumnName" :"Title",
  "ListColumnDescription" : "this is the new description"
}
```

---

## UpdateIAGExtensionLogItem


### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
    "item": 
        {
            "Title": "Test from function",
            "siteUrl": "/teams/ANL10636",
            "HubId": "b5f6c8e3-4f3e-4c2b-8f7d-1a2b3c4d5e6f",
            "Status": "Started",
            "Message": "Starting the process"
            
        }
}
```

---

## UpdateCalculatedFieldFormula


### 📖 Description

Updates the formula for a calculated field identified by SiteColumnInternalName in the
specified ListName and validates that the field exists on the specified ContentTypeName.

This function supports Managed Identity and App Only connections.

HttpStatusCode 200 : Formula successfully updated
HttpStatusCode 400 : Missing required parameters
HttpStatusCode 404 : List/content type/field not found
HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://fmsgsdevsite.sharepoint.com/sites/SPS-ae5",
  "SiteColumnInternalName": "PP_TaskReleaseSuccess",
  "NewFormula": "IF([Spec released RFC]<[Deadline Spec RFC],\"Success\",\"Failed\")",
  "ListName": "Long lead items",
  "ContentTypeName": "PP_Task"
}
```

---

## TeamifyMicrosoft365Group

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Adds Microsoft Teams capabilities to an existing Microsoft 365 Group using Add-PnPTeamsTeam.

### 💡 Examples

**Example 1**
```powershell
{
    "groupId": "11111111-2222-3333-4444-555555555555"
}
```

---

## SetTeamsChannelSiteOwners


### 📖 Description

Reads the owners of a Microsoft 365 Group and assigns them as Site Collection Administrators 
    for the associated Teams channel SharePoint site. This ensures that group owners have full 
    administrative access to manage the channel's site.
    
    HttpStatusCode 200 : Owners successfully assigned as site administrators
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error (e.g., group not found, no owners)

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "DisplayName": "test",
  "TeamsChannelUrl":"@{body('Parse_JSON')?['TeamsChannelUrl']}"
}
```

---

## SetStorageLimit


### 📖 Description

Configures the warning and maximum storage limits (in MB) for a SharePoint site collection 
    associated with a Microsoft 365 Group. Requires manual storage management to be enabled 
    in the SharePoint Admin Center settings.
    
    HttpStatusCode 200 : Storage limits successfully updated
    HttpStatusCode 400 : Invalid parameters (e.g., warning limit greater than maximum limit)
    HttpStatusCode 500 : Internal server error
    
    See https://www.sharepointdiary.com/2019/07/sharepoint-online-managing-site-storage-limits.html for more info

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "siteName": "@{body('Parse_JSON')?['siteName']}",
  "WarningLimit": "8000",
  "MaximumLimit": "10000"
}
```

---

## SetSPGroupNames

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl" : "https://contoso.sharepoint.com/sites/contoso"
  
}
```

---

## SetSPFieldProperties

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
    "siteUrl": "https://contoso.sharepoint.com/sites/contoso",
    "FieldsToUpdate": [
        {
            "FieldInternalName": "bdkClassification",
            "ListName": "Correspondence",
            "FieldProperties": [
                
                {
                    "propname": "Required",
                    "propvalue": 0
                },
                {
                    "propname": "ShowInNewForm",
                    "propvalue": "false"
                }
                    ,
                {
                    "propname": "ShowInEditForm",
                    "propvalue": "false"
                },
                {
                    "propname": "ShowInDisplayForm",
                    "propvalue": "false"
                }
            ]
        }
    ]
}
```

---

## SetSPAccessRequest

**Version:** 1.0 | 
### 📖 Description

Disables access request functionality for a SharePoint site by clearing the request access email, 
    disabling the use of access request defaults, preventing members from sharing, and disabling 
    members' ability to edit membership. Includes retry logic for reliability.
    
    HttpStatusCode 200 : Access request settings successfully disabled
    HttpStatusCode 400 : Missing required siteUrl parameter
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl" : "https://contoso.sharepoint.com/sites/contoso"   
}
```

---

## SetSiteVisitors

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Adds one or more user accounts to the Visitors group of a SharePoint site. Users are specified 
    as semicolon-separated email addresses or login names. Supports adding special groups like 
    "everyone except external users" using claims-based identities.

    HttpStatusCode 200 : All visitors successfully added
    HttpStatusCode 208 : One or more accounts could not be found or assigned
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "groupId",
  "siteName": "whatever",
  "Visitors": "AdeleV@tcwlv.onmicrosoft.com;AlexW@tcwlv.onmicrosoft.com;GradyA@tcwlv.onmicrosoft.com"
   "siteUrl": "https://contoso.sharepoint.com/sites/contoso"  #optional, bypasses the site from the group 
}
Exemple 2
{
  "groupId": "groupId",
  "siteName": "whatever",
  "Visitors": "c:0-.f|rolemanager|spo-grid-all-users/1412dec6-7931-49e3-bb90-2aae1cc9a1aa"  #this is everyone but external users
   "siteUrl": "https://contoso.sharepoint.com/sites/contoso"  #optional, bypasses the site from the group 
}
```

---

## SetSiteOwners

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Adds one or more user accounts to the Owners group of a SharePoint site. Users are specified 
    as semicolon-separated email addresses or login names. Supports adding special groups like 
    "everyone except external users" using claims-based identities.

    HttpStatusCode 200 : All owners successfully added
    HttpStatusCode 208 : One or more accounts could not be found or assigned
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
  "Owners": "AdeleV@tcwlv.onmicrosoft.com;AlexW@tcwlv.onmicrosoft.com;GradyA@tcwlv.onmicrosoft.com"
   "siteUrl": "https://contoso.sharepoint.com/sites/contoso"  
}
Exemple 2
{
  "Owners": "c:0-.f|rolemanager|spo-grid-all-users/1412dec6-7931-49e3-bb90-2aae1cc9a1aa"  #this is everyone but external users
   "siteUrl": "https://contoso.sharepoint.com/sites/contoso"  
}
```

---

## SetSiteMembers

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Adds one or more user accounts to the Members group of a SharePoint site. Users are specified 
    as semicolon-separated email addresses or login names. Supports adding special groups like 
    "everyone except external users" using claims-based identities.

    HttpStatusCode 200 : All members successfully added
    HttpStatusCode 208 : One or more accounts could not be found or assigned
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
  "Members": "AdeleV@tcwlv.onmicrosoft.com;AlexW@tcwlv.onmicrosoft.com;GradyA@tcwlv.onmicrosoft.com"
   "siteUrl": "https://contoso.sharepoint.com/sites/contoso"  
}
Exemple 2
{
  "Members": "c:0-.f|rolemanager|spo-grid-all-users/1412dec6-7931-49e3-bb90-2aae1cc9a1aa"  #this is everyone but external users
   "siteUrl": "https://contoso.sharepoint.com/sites/contoso"  
}
```

---

## SetSiteAdmins

**Version:** 1.0 | 
### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
   "SiteCollectionAdmins":"AdeleV@tcwlv.onmicrosoft.com;AlexW@tcwlv.onmicrosoft.com;c:0t.c|tenant|8b717b8f-4274-419e-a95f-c27246d80d65",
    "SiteUrl" :"https://tcwlv.sharepoint.com/sites/PnPModernSearch",
    "RemoveExistingAdmins": false
}
```

---

## SetRegionalSettings


### 📖 Description

Updates regional settings for a SharePoint site including locale, work hours, first day of week, 
    calendar type, and other locale-specific configurations. Settings can be customized via the 
    RegSettings parameter or use default values (LocaleId: 1033, WorkDay: 8am-5pm, etc.).
    
    HttpStatusCode 200 : Regional settings successfully updated
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
        "siteUrl": "https://contoso.sharepoint.com/sites/contoso"
    }
```

---

## SetPermissionForSPGroup

✅ Supports Managed Identity and App Only connections

### 📖 Description

Modifies the permission levels assigned to a SharePoint group by adding a new permission level 
    and optionally removing an existing one. This is useful for changing group permissions from 
    higher levels (e.g., Edit) to more restrictive ones (e.g., Contribute without Delete).
    
    HttpStatusCode 200 : Permissions successfully updated
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
"siteUrl":"url",
"groupName":"Member",
"addPermissionName":"Contribute without Delete",
"removePermissionName":"Edit"
}
```

---

## SetNavigation

**Version:** 1.0 | 
### 📖 Description

Adds navigation links to a SharePoint site's QuickLaunch menu. Can optionally clear existing 
    navigation before adding new nodes, and enable/disable the QuickLaunch feature. Supports 
    adding multiple navigation nodes with custom titles and URLs.
    
    HttpStatusCode 200 : Navigation successfully configured
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
    
    "ClearBeforeAdding":1,
    "EnableQuickLaunch":1, #optional
    "NavigationNodes": [
        {
            "Title": "Aktionsliste",
            "Url": "Lists/Aktionsliste",
            "Location": "QuickLaunch"
        },
        {
            "Title": "Grænsefladeaftaler",
            "Url": "Grænsefladeaftaler",
            "Location": "QuickLaunch"
        }
    ],
    "siteUrl": "https://banedanmarkonline.sharepoint.com/teams/ANL10636"
}
```

---

## SetMembersToContribute

✅ Supports Managed Identity and App Only connections

### 📖 Description

Changes the default permission level for the site's Members group from Edit to Contribute. 
    This function retrieves the SharePoint site associated with a Microsoft 365 Group, identifies 
    the Members group, and updates its permissions to the more restrictive Contribute role. 
    This is useful for reducing permissions when you want members to be able to add, edit, and 
    delete items but not manage lists or site settings.
    
    HttpStatusCode 200 : Permissions successfully updated
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
        "siteUrl": "https://contoso.sharepoint.com/sites/contoso"
    }
```

---

## SetM365SMTPEmail

✅ Supports Managed Identity and App Only connections

### 📖 Description

Change the nick name of Microsoft 365 Group by changing the email alias

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "groupId",
  "EmailDomain": "ekj.dk",
  "oldemail" : "Azure-undefined-000027@tcwlv.onmicrosoft.com",
   "tenant" : "tcwlv.onmicrosoft.com"
}
```

---

## SetM365EmailProperties

✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "EmailDomain": "ekj.dk",
  "HideFromOutlookClients" : 0
  "HideFromAddressLists" : 0
}
```

---

## SetLogo

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Configures the site logo and thumbnail for a modern SharePoint site. If only logoUrl is provided, 
    it will be used for both the site logo and thumbnail. Uses Set-PnPWebHeader to apply the images.
    
    HttpStatusCode 200 : Logo successfully applied
    HttpStatusCode 400 : Missing required logoUrl or siteUrl parameter
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
    "logoUrl":"https://banedanmarkonline.sharepoint.com/sites/itworkspaces/Style%20Library/Workspaces/90x90_kroneblack.png",
        SiteThumbnailUrl: "https://banedanmarkonline.sharepoint.com/sites/itworkspaces/Style%20Library/Workspaces/90x90_kroneblack.png",
    "siteUrl": "https://banedanmarkonline.sharepoint.com/sites/ST009521"
}
```

---

## SetLibraryVersionLimit

✅ Supports Managed Identity and App Only connections

### 📖 Description

Configures the maximum number of major versions to retain for documents in a SharePoint library. 
    This helps manage storage by limiting version history. The library name and version limit are 
    specified as a semicolon-separated pair (e.g., "Shared Documents;100").
    
    HttpStatusCode 200 : Version limit successfully applied
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
  "ListNameAndNumber": "Shared Documents;100",
  "groupId": "6ea81a76-ab55-4aa2-9e5e-4c4d34ee5b58",
  "siteUrl": "https://contoso.sharepoint.com/sites/Collab-000005" #optional bypasses the site from the group
}
```

---

## GetPnPConnection

**Version:** 1.0 | 
### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl" :"https://contoso.sharepoint.com/sites/contoso"
}
```

---

## UpdateMultiLineFieldToUseEnhancedRichText

**Version:** 1.0 | 
### 📖 Description

Finds a list column on the specified site by internal name or display title and updates
it from a regular multi-line text field to Enhanced Rich Text (FullHtml).

This function supports Managed Identity and App Only connections.

### 💡 Examples

**Example 1**
```powershell
{
    "ListName": "Long lead items",
    "FieldName": "Comments",
    "siteUrl": "https://fmsgsdevsite.sharepoint.com/sites/SPS-ae7"
}
```

---

## GetIAGMetadata

**Version:** 1.0 | 
### 📖 Description

The InternalNamesDictionary is a JSON structure that maps GUIDs to their corresponding internal names.

### 💡 Examples

**Example 1**
```powershell
{ 
  "siteId": "99f74374-63fd-4aff-9579-d105ce3f0a8d",
  OR 
    "groupId": "99f74374-63fd-4aff-9579-d105ce3f0a8d"
}
```

---

## EviDokCopyStandardDocsIntoFolders

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Traverses the entire folder structure in the source library and copies all files to the destination library.
    Creates any missing folders in the destination before copying files.

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://tenant.sharepoint.com/sites/DestinationSite",
  "sourceSiteUrl": "https://tenant.sharepoint.com/sites/SourceSite",
  "libraryName": "Documents"
}
```

---

## ApplyFormCustomizer

**Version:** 1.0 | 
### 📖 Description

Applies a JSON-based form customizer to a specific content type in a SharePoint list. 
    The JSON formatting file is retrieved from a specified SharePoint location and applied 
    to customize the form appearance and behavior.
    
    HttpStatusCode 200 : Form customizer successfully applied
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "99f74374-63fd-4aff-9579-d105ce3f0a8d",
  "listName": "Custom tab",
  "ContenttypeName": "Custom",
  "JSONFileName" :"/sites/itworkspacesadmin/Style Library/Workspaces/CustomFormatter.json",
  "JSONFileSiteUrl" : "https://banedanmarkonline.sharepoint.com/sites/itworkspacesadmin"
  
}
```

---

## ApplyAutoFillColumnsSettings

**Version:** 1.0 | 
### 📖 Description

The InternalNamesDictionary is a JSON structure that maps GUIDs to their corresponding internal names.

### 💡 Examples

**Example 1**
```powershell
{ 
  "siteUrl": "https://contoso.sharepoint.com/sites/contoso",
  "ListName": "Documents",
  "ConfigData": [
    {
      "InternalName": "AutoFill_Text",
      "IsSiteColumn": false,
      "AutofillInfo": {""LLM"":{""IsEnabled"":true,""Prompt"":""extract the name of the customer in the document. It is right after the word Customer"",""CustomModelId"":null,""CustomParametersJson"":null,""AnalyzeImageWithVision"":false,""AnalyzeImageDetailLevel"":null,""AutofillColumnType"":""Unknown""},""PrebuiltModel"":null}
    },
    {
      "InternalName": "AutoFill_Location",
      "IsSiteColumn": true,
      "AutofillInfo": {""LLM"":{""IsEnabled"":true,""Prompt"":""extract the location of the invoice, it is usually in the top right hand corner, match the location to the terms in the termset and set the term id"",""CustomModelId"":null,""CustomParametersJson"":null,""AnalyzeImageWithVision"":false,""AnalyzeImageDetailLevel"":null,""AutofillColumnType"":""Unknown""},""PrebuiltModel"":null}
    }
  ]
}
```

---

## AddTerm

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "termGroupID": "23947be5-0d03-40e3-b1a0-0b87dc3183c2",
    "termSetID": "63d628bb-2cc5-4900-86a6-e8c42e3b9554",
    "termLabel": "HR3"
}
```

---

## AddTeamsChannel

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

AddTeamsChannel is a function that will add a channel to a team.
HttpStatusCode 200 : the channel has been created
HttpStatusCode 208 : the channels exists allready

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "siteName": "@{body('Parse_JSON')?['siteName']}",
  "DisplayName": "test",
  "ChannelType" : "Standard/Shared/Private",
  "OwnerUPN": "admin@contoso.com",  #only required for private and shared channels
  "Description": "test / optional",
}
```

---

## AddSPGroupsWithPermissions

✅ Supports Managed Identity and App Only connections

### 📖 Description

Creates one or more new SharePoint groups on a site and assigns the specified permission 
    levels to each group. The permission levels must already exist in the site collection. 
    Created groups are set as members of the site's owner group.
    
    HttpStatusCode 200 : Groups successfully created with permissions
    HttpStatusCode 400 : Missing required parameters or permission level doesn't exist
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{

"siteUrl":"https://banedanmarkonline.sharepoint.com/teams/ANL10636",
      "GroupsWithPermssions": [
        {
          "GroupName" : "Byggeledelse",
          "PermissionLevelName": "Bidrag minus slet ekstern"    
        },
        {
          "GroupName" : "Assessor",
          "PermissionLevelName": "Må bidrage"
        }
        ]
}
```

---

## AddPlannerPlan

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

A longer description of the function, its purpose, common use cases, etc.

### 💡 Examples

**Example 1**
```powershell
{
    "groupId": "groupId",
    "siteName": "siteName",
    "PlannerName": "Planner",
    "Buckets": [
        {
            "Name": "Bucket1"
        },
        {
            "Name": "Bucket2"
        }
    ]

}
```

---

## AddNewFieldFromXML


### 📖 Description

AddNewFieldFromXML wraps Add-PnPFieldFromXml and is intended for adding fields,
especially calculated fields, to lists, libraries, or as site columns.

This function supports Managed Identity and App Only connections.

HttpStatusCode 200 : Field successfully added
HttpStatusCode 208 : Field already exists
HttpStatusCode 400 : Missing required parameters or invalid FieldXml
HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
Add field to a list/library
{
    "siteUrl": "https://contoso.sharepoint.com/sites/project-x",
    "ListName": "Shared Documents",
    "FieldXml": "<Field ID='{62d09286-f0bc-45ab-bf14-111ca63627f5}' Type='Calculated' Name='PP_TaskReleaseSuccess' StaticName='PP_TaskReleaseSuccess' DisplayName='Release Success' ResultType='Text'><Formula><![CDATA[=IF(AND([Spec released RFC]<[Deadline Spec RFC],[Spec released RFC]>DATE(2009,6,9)),'Success','')]]></Formula></Field>"
}
```

**Example 2**
```powershell
Add site column from XML (no ListName)
{
    "siteUrl": "https://contoso.sharepoint.com/sites/project-x",
    "FieldXml": "<Field Type='Text' Name='ProjectCode' DisplayName='Project Code' Group='Custom Columns' />"
}
```

---

## AddM365GroupOwner

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

HttpStatusCode 200 : the user has been added
HttpStatusCode 208 : the user exists allready

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "OwnerUPN": "admin@contoso.com",  #only required for private and shared channels
}
```

---

## AddIAGExtensionLogItem


### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
    "item": 
        {
            "Title": "Test from function",
            "siteUrl": "https://banedanmarkonline.sharepoint.com/teams/ANL10636",
            "HubId": "b5f6c8e3-4f3e-4c2b-8f7d-1a2b3c4d5e6f",
            "Status": "Started",
            "Message": "Starting the process",
            "Template": "Template from function"
        }
    
}
```

---

## ApplyPnPTemplate

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
    "PnPTemplateURL":"https://tcwlv.sharepoint.com/sites/SampleTeamSite/Shared%20Documents/basetemplate2.pnp",
    "siteUrl":"the site url" 
}
```

---

## AddGroupOwnersAsChannelOwners

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "@{body('Parse_JSON')?['groupId']}",
  "ChannelDisplayName": "test"
}
```

---

## AddFoldersToLibraryInBatch -deprecated

⚠️ **DEPRECATED**

**Version:** 1.1 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

⚠️ **DEPRECATED** - This function is deprecated and should not be used in new implementations.

AddFoldersToLibrary is a function that will add folders to a library based on a json file. 
Color of the folders will be set as well.
The OffsetFolder parameter identifies the folder in which the folders will be created. If the parameter is not provided, the folders will be created in the root of the library

### 💡 Examples

**Example 1**
```powershell
{
  "jsonFileSiteUrl": "https://tcwlv.sharepoint.com/sites/SampleTeamSite/",
  "ListName": "Shared Documents",
  "documentLibrary": "Documents",
  "jsonUrl": "/sites/SampleTeamSite/Shared%20Documents/FolderOptionA.json",
  "OffsetFolder" : "General" #optional
  "siteUrl": "siteurl" 
}
```

---

## AddFoldersToLibrary

**Version:** 1.1 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

AddFoldersToLibrary is a function that will add folders to a library based on a json file. 
Color of the folders will be set as well.
The OffsetFolder parameter identifies the folder in which the folders will be created. If the parameter is not provided, the folders will be created in the root of the library

### 💡 Examples

**Example 1**
```powershell
{
  "jsonFileSiteUrl": "https://tcwlv.sharepoint.com/sites/SampleTeamSite/",
  "ListName": "Shared Documents",
  "documentLibrary": "Documents",
  "jsonUrl": "/sites/SampleTeamSite/Shared%20Documents/FolderOptionA.json",
  "OffsetFolder" : "General" #optional
  "siteUrl": "siteurl" 
}
```

---

## AddDocumentTemplateToLibrary

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://tcwlv.sharepoint.com/sites/test1905",
  "templateFileUrl": "https://tcwlv.sharepoint.com/sites/RandersKommune/SiteAssets/PnP Modern Search Q.docx",
  "templateId": "PnP Modern Search Q",
  "libraryName": "Shared Documents",
  "targetContentType": "Document"
}
```

---

## AddCustomPermissionLevel

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "Custom tab",
  "BaseRoleName": "Contribute",
  "NewRoleName":"Contribute without Delete",
  "NewRoleDescription":"Contribute without Delete",
  "excludedPermissions" :  "DeleteListItems, DeleteVersions"  
}
OR
{
  
  "siteUrl": "https://banedanmarkonline.sharepoint.com/teams/ANL10619",
  "BaseRoleName": "Må bidrage",
  "NewRoleName":"Bidrage minus slet",
  "NewRoleDescription":"Bidrage minus slet",
  "excludedPermissions" :  "DeleteListItems, DeleteVersions"  

}
```

---

## AddCustomActionToGearmenu

**Version:** 1.0 | 
### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl":"https://contoso.sharepoint.com/sites/OP-1604C-test",
  "CustomActionUrl" : "https://banedanmarkonline.sharepoint.com/sites/itworkspaces/Pages/BanedanmarkFunctions.aspx?projectSiteUrl=https://banedanmarkonline.sharepoint.com/sites/TRA12715&wsDefinition=bdkbusinesssiteDK",
    "Scope" : "Web",
    "Label" : "BDK-funktioner"	,
    "BasePermission" : "ManageWeb",
    "InternalName" : "BDKFunction"
}
```

---

## AddColumnFormatting

**Version:** 1.0 | 
### 📖 Description

Applies custom JSON formatting to a specified column in a SharePoint list. The JSON formatting 
    file is retrieved from a SharePoint location and applied to customize the column's visual 
    appearance and behavior. Works with both site columns and list-specific columns.
    
    HttpStatusCode 200 : Column formatting successfully applied
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://tcwlv.sharepoint.com/sites/Collab-000007",
  "JsonUrl":"/sites/Collab-000007/Shared%20Documents/test.json",
  "ListName":"Shared Documents",
  "InternalFieldName":"Editor",
  "jsonSiteUrl":"https://tcwlv.sharepoint.com/sites/Collab-000007" 
}
```

---

## ActivateWebFeature

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

ActivateFeature is a function that will activate a feature in a site /web
each feature is identified by a GUID and separated by a semicolon

### 💡 Examples

**Example 1**
```powershell
{
    "siteUrl" :"https://contoso.sharepoint.com/sites/contoso",  
    "Features": "fa6a1bcc-fb4b-446b-8460-f4de5f7411d5"    
}
```

---

## ActivateFeature

**Version:** 2.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

ActivateFeature is a function that will activate a feature in a site collection
each feature is identified by a GUID and separated by a semicolon
see https://www.technologytobusiness.com/microsoft-sharepoint/sharepoint-site-feature-id-office-365 or https://sharepointdiv.wordpress.com/2018/04/09/sharepoint-otb-features-with-guid/
for a list of features

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl" :"https://tcwlv.sharepoint.com/sites/bdkorganisation-15007",  
  "Features": "3bae86a2-776d-499d-9db8-fa4cdc7884f8;8a4b8de2-6fd8-41e9-923c-c7c3c00f8295;b50e3104-6812-424f-a011-cc90e6327318",
}
```

---

## aa_TestConnection

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{

    "siteUrl" :"https://contoso.sharepoint.com/sites/contoso",  
    "userAssignedManagedIdentityClientId": "b50e3104-6812-424f-a011-cc90e6327318" 
}
```

---

## AddFunctionToGearmenu -deprecated

⚠️ **DEPRECATED**

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

⚠️ **DEPRECATED** - This function is deprecated and should not be used in new implementations.

A longer description of the function, its purpose, common use cases, etc.

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl":"https://contoso.sharepoint.com/sites/OP-1604C-test" 
    "Scope" : "Web",
    "Label" : "BDK-funktioner"	,
    "BasePermission" : "ManageWeb",
    "InternalName" : "BDKFunction"
}
```

---

## ApplyPnPTemplateFromXML

**Version:** 1.0 | 
### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
    "PnPXMLTempleatePath": "/pnpprovisioningfiles/lists/Correspondence.xml",
    "PnPXMLTempleateSitePath": "https://fmdkbdkdev01.sharepoint.com/sites/spoiagadminsite",
    "siteUrl": "https://fmdkbdkdev01.sharepoint.com/sites/SAM-10003-PRJ",
    "Parameters": {        
            "IAGxListTitle": "Projects",
            "IAGxListUrl": "ProjectsLinks"
        }
}
```

---

## ApplySiteMetadata-deprecated

⚠️ **DEPRECATED**

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

⚠️ **DEPRECATED** - This function is deprecated and should not be used in new implementations.

This function will apply the metadata to the site by adding the metadata to the property bag of the site

### 💡 Examples

**Example 1**
```powershell
Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional
```

**Example 2**
```powershell
lines
```

---

## ApplyTheme

**Version:** 1.1 | 
### 📖 Description

ApplyTheme is a function that will apply a modern theme to a site collection.
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
   "Themename" : "Red",
  "siteUrl": "url" #optional
}
```

---

## EviDocSetColumnDefaultValues

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

Traverses the folder structure under the given first-level folder and sets the default column value
    for EviDokFolderAsString to the pipe-separated folder path (e.g. FolderA|SubFolder|DeepFolder).

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://tenant.sharepoint.com/sites/DestinationSite",
  "libraryName": "Documents",
  "firstLevelFolderName": "FolderA"
}
```

---

## EviDocCopyFolders

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "https://tenant.sharepoint.com/sites/DestinationSite",
  "sourceSiteUrl": "https://tenant.sharepoint.com/sites/SourceSite",
  "libraryName": "Documents",
  "firstLevelFolderName": "FolderA",
  colorHex: 11 #optional, default is 11 (Green)
}
```

---

## EnableDisableGridEditingOnList

**Version:** 1.0 | 
### 📖 Description

Controls the grid editing feature for a SharePoint list. Grid editing allows users to quickly 
    edit multiple list items in a spreadsheet-like interface. Use DisableGridEditing parameter 
    set to 0 to enable or 1 to disable.
    
    HttpStatusCode 200 : Grid editing setting successfully updated
    HttpStatusCode 400 : Missing required parameters
    HttpStatusCode 500 : Internal server error
    
    Note: DisableGridEditing isn't in the PnP Provisioning Schema yet.
    See https://github.com/pnp/PnP-Provisioning-Schema/issues/607
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "url",
  "ListName": "listname",
  "DisableGridEditing": 0/1
}
```

---

## DisableWebTemplatesGalleryFirstRunDialog

**Version:** 1.0 | ✅ Supports Managed Identity and App Only connections

### 📖 Description

This Azure Function will disable the Web Template Gallery picker that pops up the first time the newly created site is opened.

### 💡 Examples

**Example 1**
```powershell
{
    "siteUrl":"https://contoso.sharepoint.com/sites/OP-1604C-test"
}
```

---

## DisableSocialBarOnSitePages

✅ Supports Managed Identity and App Only connections

### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{   
    "SiteUrl" : "https://contoso.sharepoint.com/sites/contoso"    
    }
```

---

## DisableNextStepsFirstRunEnabled

**Version:** 1.0 | 
### 📖 Description

Disables the "Next Steps" first run dialog that appears to users when they first visit a 
    SharePoint site. This provides a cleaner initial experience for sites that don't need the 
    guided setup process.
    
    HttpStatusCode 200 : Next Steps first run successfully disabled
    HttpStatusCode 400 : Missing required siteUrl parameter
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "url"
}
```

---

## DisableListComments

✅ Supports Managed Identity and App Only connections

### 📖 Description

The “Comment” feature introduced in SharePoint Online modern lists allows users to comment on SharePoint Online list items. 
    Commenting is a great way to collect feedback and keep track of changes in essential documents. 
    However, in some cases, you may want to disable comments on a list to prevent unwanted content

### 💡 Examples

**Example 1**
```powershell
{   
    "SiteUrl" : "https://contoso.sharepoint.com/sites/contoso",
    "ListName" : "Ændringslog Entreprenør"
    }
```

---

## DeleteViewsFromLists

**Version:** 1.0 | 
### 📖 Description

Reviews named view from named lists
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "url",
    "ToDeDeletedViewsInLists": [
  {
    "ListName": "ListName",
    "ViewName": "ViewName"
  },

  
}
```

---

## DeleteList

**Version:** 1.0 | 
### 📖 Description

Deletes a named list from a site.
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  
  "ListName" : "Documents",
  "siteUrl": "the site url" 
}
```

---

## DeleteContentTypesFromList

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl": "url",
  "ToDeDeletedContentTypes": [
  {
    "ListName": "ListName",
    "ContentTypeName": "ContentTypeName"
  },
  {
    "ListName": "ListName",
    "ContentTypeName": "ContentTypeName"
  }
  ]
}
```

---

## DeleteContentTypeFromLibrary-deprecated

⚠️ **DEPRECATED**

✅ Supports Managed Identity and App Only connections

### 📖 Description

⚠️ **DEPRECATED** - This function is deprecated and should not be used in new implementations.

This Azure Function will delete content types from a list
THis is often useful when you add another Content type to a list and you want to remove the old one

### 💡 Examples

**Example 1**
```powershell
{
  "groupId": "77e4df47-2c32-4c7b-a6a3-beca9b674e07",
  "siteName": "test",
    "ListName" : "Documents",
  "ContentTypesToBeDeleted":"ISO 14001 Policy",
  "siteUrl":"the Url" #optional
}
```

---

## CreateM365TeamsTab_WebSite

**Version:** 1.0 | 
### 📖 Description

Creates a new tab in a Microsoft Teams channel that displays an external website. The tab 
    is configured with the specified website URL (contentUrl) and will be visible to all 
    channel members.
    
    HttpStatusCode 200 : Tab successfully created
    HttpStatusCode 400 : Missing required parameters or invalid tabType
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "channelName": "General",
  "contentUrl": "https://www.dr.dk",
  "groupId": "99f74374-63fd-4aff-9579-d105ce3f0a8d",
  "tabName": "website tab",
  "tabType": "WebSite",
  "siteUrl": "https://tcwlv.sharepoint.com/sites/Azure-undefined-000026-testB"
}
```

---

## CreateM365TeamsTab_SharePointPageAndList

**Version:** 1.1 | 
### 📖 Description

Creates a new tab in a Microsoft Teams channel that displays either a SharePoint page or list. 
    Can point to the default frontpage or a specific page. Requires the SharePoint site URL, 
    list name, and page content URL.
    
    HttpStatusCode 200 : Tab successfully created
    HttpStatusCode 400 : Missing required parameters or invalid tabType
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
1  default frontpage
{
  "channelName": "General", 
  "groupId": "99f74374-63fd-4aff-9579-d105ce3f0a8d",
  "tabName": "test",
  "tabType": "SharePointPageAndList",
  "siteUrl": "https://tcwlv.sharepoint.com/sites/Azure-undefined-000025-testB/",
  "ListName": "something",
  "pageContentUrl": "https://tcwlv.sharepoint.com/sites/Azure-undefined-000026/"  
}
```

**Example 2**
```powershell
1  specific page
{
  "channelName": "General", 
  "groupId": "99f74374-63fd-4aff-9579-d105ce3f0a8d",
  "tabName": "test",
  "tabType": "SharePointPageAndList",
  "siteUrl": "https://tcwlv.sharepoint.com/sites/Azure-undefined-000025-testB/",
  "ListName": "something",
  "pageContentUrl": "https://tcwlv.sharepoint.com/sites/Azure-undefined-000026/SitePages/Home.aspx"  
}
```

---

## CreateM365TeamsTab_Planner

**Version:** 1.0 | 
### 📖 Description

Creates a new tab in a Microsoft Teams channel that displays a Microsoft Planner plan. 
    Requires the planner plan name and tenant name to locate and link the correct plan. 
    This enables team members to view and manage tasks directly within Teams.
    
    HttpStatusCode 200 : Tab successfully created
    HttpStatusCode 400 : Missing required parameters or invalid tabType
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
"groupId" :"add6c948-e3e7-4570-8d08-fcf44845fe38",
"channelName" :"General",
"tabName":"Test Library link",
"tabType":"Planner",
"PlannerName" :"My Planner" #the name of the planner plan
"TenantName": "tcwlv.onmicrosoft.com"
}
```

---

## CreateM365TeamsTab_OneNote

**Version:** 1.0 | 
### 📖 Description

Creates a new tab in a Microsoft Teams channel that displays the team's OneNote notebook. 
    This provides quick access to shared notes and collaborative documentation within the 
    Teams interface.
    
    HttpStatusCode 200 : Tab successfully created
    HttpStatusCode 400 : Missing required parameters or invalid tabType
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
"groupId" :"add6c948-e3e7-4570-8d08-fcf44845fe38",
"channelName" :"General",
"tabName":"Test Library link",
"tabType":"OneNote"
}
```

---

## CreateM365TeamsTab_DocumentLibrary

**Version:** 1.0 | 
### 📖 Description

Creates a new tab in a Microsoft Teams channel that displays a SharePoint document library. 
    The tab provides direct access to the specified library within the Teams interface, making 
    it easy for team members to access and collaborate on documents.
    
    HttpStatusCode 200 : Tab successfully created
    HttpStatusCode 400 : Missing required parameters or invalid tabType
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
1
{
"groupId" :"add6c948-e3e7-4570-8d08-fcf44845fe38",
"channelName" :"General",
"tabName":"Test Library link",
"tabType":"DocumentLibrary",
"siteUrl":"https://tcwlv.sharepoint.com/sites/KOMBITTARGETSITE",
"targetLibraryUrl":"https://tcwlv.sharepoint.com/sites/KOMBITTARGETSITE/Shared%20Documents"
}
```

---

## CreateM365TeamsTab_Custom

**Version:** 1.0 | 
### 📖 Description

Creates a new tab in a Microsoft Teams channel using a custom Teams app. Requires the 
    App ID of the custom Teams application to be installed in the team. This enables integration 
    of custom business applications within the Teams interface.
    
    HttpStatusCode 200 : Tab successfully created
    HttpStatusCode 400 : Missing required parameters or invalid tabType
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "channelName": "General",
  "AppId": "f10c0d2f-c73c-4a79-a351-28392fc440db",
  "groupId": "99f74374-63fd-4aff-9579-d105ce3f0a8d",
  "tabName": "Custom tab",
  "tabType": "Custom"
  
}
```

---

## CreateM365TeamsTab


### 📖 Description

Creates a new tab in a Microsoft Teams channel. Supports multiple tab types including 
    DocumentLibrary, SharePointPageAndList, OneNote, Planner, and WebSite. This is a generic 
    function that routes to the appropriate specialized tab creation logic based on the tabType parameter.
    
    HttpStatusCode 200 : Tab successfully created
    HttpStatusCode 400 : Missing required parameters or invalid tabType
    HttpStatusCode 500 : Internal server error
    
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
1
{
"groupId" :"add6c948-e3e7-4570-8d08-fcf44845fe38",
"channelName" :"General",
"siteUrl": "https://tcwlv.sharepoint.com/sites/Azure-undefined-000025-testB/",  #for lists and libraries
"tabName":"Test Library link",
"tabType":"DocumentLibrary",
"targetLibraryUrl":"https://tcwlv.sharepoint.com/sites/KOMBITTARGETSITE/Shared%20Documents"
OR 
contentUrl = "https://www.fellowmind.com/"
}
```

**Example 2**
```powershell
2
{
"groupId" :"add6c948-e3e7-4570-8d08-fcf44845fe38",
"channelName" :"General",
"tabName":"Test Library link",
"tabType":"OneNote"
}
```

**Example 3**
```powershell
3
{
"groupId" :"add6c948-e3e7-4570-8d08-fcf44845fe38",
"channelName" :"General",
"tabName":"Test Library link",
"tabType":"Planner",
"PlannerName" :"My Planner" #the name of the planner plan
"TenantName": "tcwlv.onmicrosoft.com"

}
```

**Example 4**
```powershell
4
{
  "channelName": "testB",
  "groupId": "b434a167-2dc6-420f-92e1-fa5701998a0c",
  "tabName": "local Library link",
   "tabType": "SharePointPageAndList",
  "ListName": "Lists/tomliste",
"siteUrl": "https://tcwlv.sharepoint.com/sites/Azure-undefined-000025-testB/"

}
```

---

## CopyCustomPageTemplates

**Version:** 1.0 | 
### 📖 Description

CopyCustomPageTemplates is a function that will copy custom page templates from a source site to a destination site.
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
    "siteUrl": "https://contoso.sharepoint.com/sites/DestinationSite",
    "SourceUrl": "https://contoso.sharepoint.com/sites/SourceSite"
}
```

---

## AreContentTypesFromHubReady


### 📖 Description



### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl" :"https://contoso.sharepoint.com/sites/contoso",  
  "ContentTypes": "0x0101;0x01"
}
```

---

## ArchiveSiteUsingMSArchive

**Version:** 1.0 | 
### 📖 Description

is a function that will [do somehting]
This function supports Managed Identity and App Only connections

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl" : "https://contoso.sharepoint.com/sites/Custom" 
}
```

---

## GetContentTypesFromHub

✅ Supports Managed Identity and App Only connections

### 📖 Description

ActivateFeature is a function that will activate a feature in a site collection
each feature is identified by a GUID and separated by a semicolon
see https://www.technologytobusiness.com/microsoft-sharepoint/sharepoint-site-feature-id-office-365 or https://sharepointdiv.wordpress.com/2018/04/09/sharepoint-otb-features-with-guid/
for a list of features

### 💡 Examples

**Example 1**
```powershell
{
  "siteUrl" :"https://contoso.sharepoint.com/sites/contoso",  
  "ContentTypes": "0x0101;0x01"
}
```

---

## UpdatePortalItem


### 📖 Description

This function updates a specific field in a listitem in a list.

### 💡 Examples

**Example 1**
```powershell
Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional
```

**Example 2**
```powershell
lines
```


