Apache Cordova plugin for Microsoft Azure Active Directory Graph
=============================
Provides JavaScript API to work with Microsoft Azure Active Directory Graph.
####Supported Platforms####

- Android (cordova-android@>=4.0.0 is supported)
- iOS
- Windows (Windows 8.0, Windows 8.1 and Windows Phone 8.1)

## Sample usage ##
To access the [AAD Graph API](https://msdn.microsoft.com/en-us/library/azure/hh974476.aspx) you need to acquire an access token and get the AAD client. Then, you can send async queries to interact with AAD entities. Note: application ID, authorization and redirect URIs are assigned when you register your app with Microsoft Azure Active Directory.

```javascript
var resourceUrl = 'https://graph.windows.net/';
var tenantName = 'sampleDirectory2015.onmicrosoft.com';
var endpointUrl = resourceUrl + tenantName;
var appId = '98ba0820-f7da-4411-87bb-598e0475536b';
var authority = 'https://login.windows.net/' + tenantName + '/';
var redirectUrl = 'http://localhost:4400/services/aad/redirectTarget.html';

var AuthenticationContext = Microsoft.ADAL.AuthenticationContext;
var ActiveDirectoryClient = Microsoft.AADGraph.ActiveDirectoryClient;

var authContext = new AuthenticationContext(authority);
var client = new ActiveDirectoryClient(endpointUrl, authContext, resourceUrl, appId, redirectUrl);

client.users.getUsers().fetchAll().then(function (users) {
    result.forEach(function (user) {
        console.log('User: ' + user.displayName + ' PrincipalName: ' + user.userPrincipalName);
    });
}, function(error) {
    console.log(error);
});
```

Complete example is available [here](https://github.com/AzureAD/azure-activedirectory-cordova-plugin-graph/tree/master/sampleApp).

## Installation Instructions ##

Use [Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) to create your app and add the plugin.

1. Make sure an up-to-date version of Node.js is installed, then type the following command to install the [Cordova CLI](https://github.com/apache/cordova-cli):

        npm install -g cordova

2. Create a project and add the platforms you want to support:

        cordova create aadClientApp
        cd aadClientApp
        cordova platform add windows <- support of Windows 8.0, Windows 8.1 and Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. Add the plugin to your project:

        cordova plugin add https://github.com/AzureAD/azure-activedirectory-cordova-plugin-graph

4. Build and run, for example:

        cordova run android

To learn more, read [Apache Cordova CLI Usage Guide](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

## Common operations ##

### Getting entities ###
```javascript
// Option #1: client.<entities>.get<Entities>().fetchAll().then(win, fail);
client.users.getUsers().fetchAll().then(function(itemsArray) { 
    ... 
}, function(err) { 
    ... 
});

// Option #2: client.directoryObjects.as<Entities>().getDirectoryObjects().fetchAll().then(win, fail);
client.directoryObjects.asApplications().getDirectoryObjects().fetchAll().then(function(itemsArray) { 
    ...
}, function(err) { 
    ... 
});

// Get by objectId: client.<entities>.get<Entity>(objectId).fetch().then(win, fail);
client.devices.getDevice('33f8f7d8-38e4-44f8-acf3-2745cb6d2e80').fetch().then(function(device) { 
    ... 
}, function(err) { 
    ... 
});

// Or alternatively:
client.directoryObjects.asUsers().getDirectoryObject('33f8f7d8-38e4-44f8-acf3-2745cb6d2e80').fetch().then(function(user) { 
    ... 
}, function(err) { 
    ... 
});
```

### Adding entities ###
```javascript
// Option #1: client.<entities>.add<Entity>(instance).then(win, fail);
client.users.addUser(user).then(function(addedUser) { 
    ... 
}, function(err) { 
    ... 
});

// Option #2: client.directoryObjects.as<Entities>().addDirectoryObject(instance).then(win, fail);
client.directoryObjects.asGroups().addDirectoryObject(group).then(function(addedGroup) { 
    ... 
}, function(err) { 
    ... 
});
```

### Updating entities ###
```javascript
client.groups.getGroup(objectId).fetch().then(function (groupFoundById) {
    groupFoundById.displayName = 'Updated group name';
    groupFoundById.update().then(function () { ... }, function (err) { ... });
}, function (err) { ... });
```

### Deleting entities ###
```javascript
client.groups.getGroup(objectId).fetch().then(function (groupFoundById) {    
    groupFoundById.delete().then(function () { ... }, function (err) { ... });
}, function (err) { ... });
```

### [Queries, Filters and Paging](https://msdn.microsoft.com/en-us/library/azure/dn727074.aspx) ###

#### filter ####
```javascript
client.users.getUsers().filter("displayName eq 'John Trident'").fetchAll().then( ... );
client.users.getUsers().filter("startswith(displayName,'test')").fetchAll().then( ... );
```

#### top ####
```javascript
client.users.getUsers().top(3).fetchAll().then( ... )
```

#### orderBy ####
```javascript
client.users.getUsers().orderBy('displayName').fetchAll().then( ... )
```

#### Combinations ####
```javascript
client.users.getUsers().top(3).orderBy('displayName').fetchAll().then( ... )
```
__Note__: combining of `filter` and `orderBy` is not supported.

#### paging ####
Results are being paginated when `fetch` method is used:
```javascript
client.users.getUsers().fetch().then(function(pagedCollection) { 
    pagedCollection.currentPage.forEach(function (user) {
        console.log('User "' + user.displayName + '" userPrincipalName: "' + user.userPrincipalName + '"');
    });

    pagedCollection.getNextPage().then(function(secondPagedCollection) {
        ...
    }, function(err) { 
        ... 
    });
}, function(err) { 
    ... 
});
```
__Notes__: 
- Queries are not yet supported along with `fetch` by AAD Graph JS SDK,
- getPreviousPage is not yet supported by AAD Graph JS SDK.

## Entity types and their specific operations ##

### [Users](https://msdn.microsoft.com/en-us/library/azure/hh974483.aspx) ###

#### Creating and initializing an entity instance ####
```javascript
var user = new AadGraph.User();
user.displayName = 'displayName';
user.mailNickname = 'mailNickname';
user.userPrincipalName = displayName + '@' + tenantName;
var passwordProfile = new AadGraph.PasswordProfile();
passwordProfile.password = "tempPassword1234";
passwordProfile.forceChangePasswordNextLogin = true;
user.passwordProfile = passwordProfile;
// `usageLocation` property is required in order to use `assignLicense` method (see example below):
user.usageLocation = 'US';
user.accountEnabled = true;
```

#### [Managing subscriptions and plans](https://msdn.microsoft.com/en-us/library/azure/dn835115.aspx) ####

You can call on user `assignLicense` to add or remove subscriptions for the user. 
You can also enable and disable specific plans associated with a subscription.  
##### Adding a subscription with its service plans: 
```javascript
user.assignLicense([
{
    "disabledPlans": [],
    "skuId": subscribedSkuId
}], []).then(function () { ... }, function(err));
```
##### Disabling subscription plan:
```javascript
user.assignLicense([
{
    "disabledPlans": [servicePlanIdToDisable],
    "skuId": subscribedSkuId
}], []).then(function () { ... }, function(err));
```
##### Removing subscription:
```javascript
user.assignLicense([], [subscribedSkuIdToRemove]);
```

#### Reseting user password ####
To reset user password you should update its `PasswordProfile` property.
For example:
```javascript
user.passwordProfile = new AadGraph.PasswordProfile();
user.passwordProfile.password = "ChangedPass1234";
user.passwordProfile.forceChangePasswordNextLogin = true;
user.update().then( ... );
```

#### [Manager](https://msdn.microsoft.com/en-us/library/azure/dn151688.aspx) ####
```javascript
// Getting
user.manager.fetch().then(function(manager) { ... }, function(err) { ... });

// Updating
user.update_manager(anotherUser).then(function () { ... }, function(err) { ... });
```

#### [DirectReports](https://msdn.microsoft.com/en-us/library/azure/dn151686.aspx) ####
The Get User’s Direct Reports operation returns a list of the user’s direct reports. These are users and contacts that have their `manager` navigation property set to the user on which the operation is performed. 
```javascript
user.directReports.getDirectoryObjects().fetchAll().then(function(reports) { 
    ...
}, function(err) { 
    ...
});
```

#### [memberOf](https://msdn.microsoft.com/en-us/library/azure/dn151667.aspx) ####
```javascript
user.memberOf.getDirectoryObjects().fetchAll().then( ... )
```

#### [OAuth2PermissionGrants](https://msdn.microsoft.com/en-us/library/azure/dn151672.aspx) ####
```javascript
user.oauth2PermissionGrants.getOAuth2PermissionGrants().fetchAll().then( ... );
```

#### [AppRoleAssignments](https://msdn.microsoft.com/en-us/library/azure/dn835128.aspx) ####
```javascript
user.appRoleAssignments.getAppRoleAssignments().fetchAll().then( ... )

user.appRoleAssignments.addAppRoleAssignment(newItem).then(function(addedItem) { 
    ... 
}, function(err) { 
    ... 
});
```

### [Groups](https://msdn.microsoft.com/en-us/library/azure/hh974486.aspx) ###

#### Creating and initializing an entity instance ####
```javascript
var group = new AadGraph.Group(client.context);
group.displayName = 'myGroup';
group.mailNickname = 'groupMailNickName';
group.mailEnabled = 'false';
group.securityEnabled = 'true';`
```

#### Members ####
```javascript
group.members.getDirectoryObjects().fetchAll().then(function(members) { 
    ... 
}, function(err) { 
    ... 
});

group.addMember(newMember).then(function () { 
    ... 
}, function(err) { 
    ... 
});

group.deleteMember(existingMember).then(function () { 
    ... 
}, function(err) { 
    ... 
});
```
__Note__: `members` property is applicable to [DirectoryRole](https://msdn.microsoft.com/en-us/library/azure/jj134103.aspx) as well.

#### [isMemberOf](https://msdn.microsoft.com/en-us/library/azure/dn835135.aspx) ####
```javascript
client.isMemberOf(group.objectId, user.objectId).then(function(obj) {
   if (obj.value === true) { console.log('the member is in the group'); } 
}, function(err) { ... });
```

#### [getMemberGroups](https://msdn.microsoft.com/en-us/library/azure/dn835126.aspx) ####
Call the getMemberGroups function on a user, contact, group, or service principal to get the group objectId's that it is a member of.
You can pass  boolean argument to get only security enabled groups.  
For example,
```javascript
// `securityEnabledOnly` - false by default
user.getMemberGroups().then(function (groupIds) {}, function(err) { ... });

// `securityEnabledOnly` = true
user.getMemberGroups(true).then(function (groupIds) {}, function(err) { ... });
```

#### [getMemberObjects](https://msdn.microsoft.com/en-us/library/azure/dn835117.aspx) ####
This function is similar to `getMemberGroups` but it returns role objectIds as well.

#### [checkMemberGroups](https://msdn.microsoft.com/en-us/library/azure/dn835107.aspx) ####
You can call `checkMemberGroups` on an instance to check its membership in a list of groups.
```javascript
user.checkMemberGroups([group1.objectId, group2.objectId]).then(function (groupIds) {
    // groupIds is an array containing parent groups objectId's
}, function (err) { 
    ... 
});
```

### [Applications](https://msdn.microsoft.com/en-us/library/azure/dn151677.aspx) ###

#### Creating and initializing an entity instance ####
```javascript
var app = new AadGraph.Application(client.context);
var identifierUrl = 'https://contoso.com/app';
app.displayName = 'MyApp';
app.identifierUris = [identifierUrl];
```

#### Getting deleted applications ####
```javascript
client.deletedDirectoryObjects.asApplications().fetchAll().then(function(apps) { ... }, function(err) { ... })
```

#### Restoring deleted applications ####
`identifierUris` argument can be passed to `restore` method if you need to update the old value.

For example:
```javascript
client.deletedDirectoryObjects.asApplications().fetchAll().then(function (apps) {
    apps[0].restore().then(function (restoredApp) {
        apps[1].restore(['https://contoso.com/app']).then(function (secondRestoredApp) { 
            ... 
        }, function(err) { ... });
    }, function(err) { ... });
}, function(err) { ... });
```

### [Directory Schema Extensions](https://msdn.microsoft.com/en-us/library/azure/dn720459.aspx) ###

#### Creating and initializing an entity instance ####
```javascript
var property = new AadGraph.ExtensionProperty();
property.name = 'skypeId';
property.dataType = 'String';
property.targetObjects = property.targetObjects.concat('User');
```

#### Adding an extension ####
```javascript
client.applications.addApplication(that.createApp()).then(function (app) {
    app.extensionProperties.addExtensionProperty(extensionRaw).then(function (extension) {
        // The name should become a fully-qualified extension property name like 'extension_ab603c56068041afb2f6832e2a17e237_skypeId'
        console.log(extension.name);
        ...
    });
});
```

#### Writing and reading an extension property ####
```javascript
...
user['extension_ab603c56068041afb2f6832e2a17e237_skypeId'] = 'live:userSkypeId';
user.update().then(function() {
    client.users.getUser(user.objectId).fetch().then(function (updatedUser) {
        console.log(updatedUser['extension_ab603c56068041afb2f6832e2a17e237_skypeId']);
    });
});
```
__Note__: extension properties reading and writing are not currently supported by JS SDK.

### Creating and initializing other entity types ###

#### Creating a service principal for an app ####
```javascript
var principal = new AadGraph.ServicePrincipal();
principal.appId = app.appId;
```

#### Creating an oauth2PermissionGrant ####
```javascript
var grant = new AadGraph.OAuth2PermissionGrant(client.context);
grant.resourceId = resourceId;
grant.clientId = clientId;
grant.principalId = principalId;
grant.consentType = 'Principal';
grant.startTime = '2014-03-01';
grant.expiryTime = '2014-04-01';
```

#### Creating a device ####
```javascript
var device = new AadGraph.Device();
device.displayName = 'myDevice';
device.deviceId = deviceId;
device.accountEnabled = true;
var altSecId = new AadGraph.AlternativeSecurityId();
altSecId.key = btoa(key);
altSecId.type = type;
// Triggering _identityProviderChanged to become true as this field is required in request
altSecId.identityProvider = null;
device.alternativeSecurityIds = [altSecId];
device.deviceOSType = OSType;
device.deviceOSVersion = OSVersion;
```

#### Creating an appRoleAssignment ####
```javascript
var assignment = new AadGraph.AppRoleAssignment();
assignment.id = id;
assignment.resourceId = resourceId;
assignment.principalId = principalId;
```

## Copyrights ##
Copyright (c) Microsoft Open Technologies, Inc. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License"); you may not use these files except in compliance with the License. You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
