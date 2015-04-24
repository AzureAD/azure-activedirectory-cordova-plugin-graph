
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
    'use strict';

    angular
      .module('starter.services')
      .factory('AadClient', ['tenantName', 'authority', 'resourceUrl', 'redirectUrl', 'appId', 'PromiseFactory', AadClient]);

    function AadClient(tenantName, authority, resourceUrl, redirectUrl, appId, PromiseFactory) {
        var officeEndpointUrl = resourceUrl + '/' + tenantName;

        var authContext = new Microsoft.ADAL.AuthenticationContext(authority);
        var client = new Microsoft.AADGraph.ActiveDirectoryClient(officeEndpointUrl, authContext, resourceUrl, appId, redirectUrl);

        var service = {
            getUsers: getUsers,
            getUser: getUser,
            addUser: addUser,
            editUser: editUser,
            deleteUser: deleteUser,
            resetPassword: resetPassword,
            getGroups: getGroups,
            getGroup: getGroup,
            addGroup: addGroup,
            editGroup: editGroup,
            deleteGroup: deleteGroup,
            getGroupMembers: getGroupMembers,
            deleteGroupMember: deleteGroupMember,
            getPossibleGroupMembers: getPossibleGroupMembers,
            addGroupMembers: addGroupMembers,
            getApps: getApps,
            getApp: getApp,
            addApp: addApp,
            editApp: editApp,
            deleteApp: deleteApp,
            getDeletedApps: getDeletedApps,
            getDeletedApp: getDeletedApp,
            restoreApp: restoreApp,
            authenticate: authenticate,
            logOut: logOut
        };
        return service;

        //////////////

        function logError(err) {
            console.log(err + (err && err.responseText) ? (': ' + err.responseText) : '');
        }

        function onError(err) {
            logError(err);
            this.reject(err);
        }

        function getUsers() {
            var promise = PromiseFactory.createPromise();

            client.users.getUsers().fetchAll().then(function (users) {
                promise.resolve(users);
            }, onError.bind(promise));

            return promise;
        }

        function getUser(userId) {
            var promise = PromiseFactory.createPromise();

            client.users.getUser(userId).fetch().then(function (user) {
                promise.resolve(user);
            }, onError.bind(promise));

            return promise;
        }

        function addUser(userName, displayName, firstName, lastName) {
            var promise = PromiseFactory.createPromise();
            var newUser = createUser(userName, displayName, firstName, lastName);

            client.users.addUser(newUser).then(function (user) {
                user.passwordProfile = newUser.passwordProfile; //to return temporary password
                promise.resolve(user);
            }, onError.bind(promise));

            return promise;
        }

        function editUser(user, name, displayName, firstName, lastName) {
            var promise = PromiseFactory.createPromise();

            user.displayName = displayName;
            user.mailNickname = name + 'MailNickname';
            user.userPrincipalName = name + '@' + tenantName;

            user.givenName = firstName;
            user.surname = lastName;

            user.update().then(function () {
                promise.resolve();
            }, onError.bind(promise));

            return promise;
        }

        function createUser(userName, displayName, firstName, lastName) {
            var name = userName || displayName || getGuid();
            name = trimInternalSpaces(name);

            var user = new AadGraph.User();
            user.displayName = displayName;
            user.accountEnabled = true;
            user.mailNickname = name + 'MailNickname';
            user.userPrincipalName = name + '@' + tenantName;

            if (firstName) {
                user.givenName = firstName;
            }

            if (lastName) {
                user.surname = lastName;
            }

            var passwordProfile = new AadGraph.PasswordProfile();
            passwordProfile.password = generatePassword();
            passwordProfile.forceChangePasswordNextLogin = true;
            user.passwordProfile = passwordProfile;

            return user;
        }

        function generatePassword() {
            var temporaryPassword = getGuid();
            temporaryPassword = temporaryPassword.substr(0, temporaryPassword.indexOf('-'));

            //password should have the symbol in upper case
            //TODO: need to improve the password?
            temporaryPassword = "Q" + temporaryPassword.substr(1) +
                Math.floor(Math.random() * 10) + "@c";

            return temporaryPassword;
        }

        function trimInternalSpaces(name) {
            var result = '';
            for (var i = 0; i < name.length; i++) {
                if (name[i] !== ' ') {
                    result += name[i];
                }
            }

            return result;
        }

        function getGuid() {
            function _p8(s) {
                var p = (Math.random().toString(16) + "000000000").substr(2, 8);
                return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
            }
            return _p8() + _p8(true) + _p8(true) + _p8();
        }

        function deleteUser(userId) {
            var promise = PromiseFactory.createPromise();

            getUser(userId).then(function (user) {
                user.delete().then(function () {
                    promise.resolve();
                }, onError.bind(promise));
            }, onError.bind(promise));

            return promise;
        }

        function resetPassword(user) {
            var promise = PromiseFactory.createPromise();

            var passwordProfile = new AadGraph.PasswordProfile();
            passwordProfile.password = generatePassword();
            passwordProfile.forceChangePasswordNextLogin = true;
            user.passwordProfile = passwordProfile;

            user.update().then(function() {
                promise.resolve(user.passwordProfile.password);
            }, onError.bind(promise));

            return promise;
        }

        function getGroups() {
            var promise = PromiseFactory.createPromise();

            client.groups.getGroups().fetchAll().then(function (groups) {
                promise.resolve(groups);
            }, onError.bind(promise));

            return promise;
        }

        function getContacts() {
            var promise = PromiseFactory.createPromise();

            client.contacts.getContacts().fetchAll().then(function (contacts) {
                promise.resolve(contacts);
            }, onError.bind(promise));

            return promise;
        }

        function getGroup(groupId) {
            var promise = PromiseFactory.createPromise();

            client.groups.getGroup(groupId).fetch().then(function (group) {
                promise.resolve(group);
            }, onError.bind(promise));

            return promise;
        }

        function getDirectoryObject(objectId) {
            var promise = PromiseFactory.createPromise();

            client.directoryObjects.getDirectoryObject(objectId).fetch().then(function (directoryObject) {
                promise.resolve(directoryObject);
            }, onError.bind(promise));

            return promise;
        }

        function addGroup(displayName, description) {
            var promise = PromiseFactory.createPromise();
            var newGroup = createGroup(displayName, description);

            client.groups.addGroup(newGroup).then(function (group) {
                promise.resolve(group);
            }, onError.bind(promise));

            return promise;
        }

        function editGroup(group, displayName, description) {
            var promise = PromiseFactory.createPromise();

            group.displayName = displayName;
            group.description = description;

            group.update().then(function () {
                promise.resolve();
            }, onError.bind(promise));

            return promise;
        }

        function createGroup(displayName, description) {
            var group = new AadGraph.Group(AadClient.context, null, null);
            group.displayName = displayName || 'testGroup1';
            group.description = description;
            group.mailNickname = 'test';
            group.mailEnabled = 'false';
            group.securityEnabled = 'true';
            return group;
        }

        function deleteGroup(groupId) {
            var promise = PromiseFactory.createPromise();

            getGroup(groupId).then(function (group) {
                group.delete().then(function () {
                    promise.resolve();
                }, onError.bind(promise));
            }, onError.bind(promise));

            return promise;
        }

        function getGroupMembers(groupId) {
            var promise = PromiseFactory.createPromise();

            getGroup(groupId).then(function (group) {
                group.members.getDirectoryObjects().fetchAll().then(function (members) {
                    promise.resolve(members);
                }, onError.bind(promise));
            }, onError.bind(promise));

            return promise;
        }

        function deleteGroupMember(groupId, memberId) {
            var promise = PromiseFactory.createPromise();

            getGroup(groupId).then(function (group) {
                getDirectoryObject(memberId).then(function (member) {
                    group.deleteMember(member).then(function () {
                        promise.resolve();
                    }, onError.bind(promise));
                }, onError.bind(promise));
            }, onError.bind(promise));

            return promise;
        }

        function getPossibleGroupMembers(groupId) {
            var promise = PromiseFactory.createPromise();

            // TODO: refactor this
            //function checkDupAmongExistingMembers(existingMember) {
            //    return this.objectId !== existingMember.objectId;
            //}

            getGroup(groupId).then(function (group) {
                group.members.getDirectoryObjects().fetchAll().then(function (members) {
                    getGroups().then(function (allGroups) {
                        getUsers().then(function (allUsers) {
                            getContacts().then(function (allContacts) {
                                promise.resolve(allGroups.filter(function (item) {
                                    return item.objectId !== group.objectId
                                        && members.filter(function (existingMember) {
                                            return existingMember.objectId === item.objectId;
                                        }).length === 0;
                                }).concat(allUsers.filter(function (user) {
                                    return members.filter(function (existingMember) {
                                        return existingMember.objectId === user.objectId;
                                    }).length === 0;
                                }), allContacts.filter(function (contact) {
                                    return members.filter(function (existingMember) {
                                        return existingMember.objectId === contact.objectId;
                                    }).length === 0;
                                })));
                            }, onError.bind(promise));
                        }, onError.bind(promise));
                    }, onError.bind(promise));
                }, onError.bind(promise));
            }, onError.bind(promise));

            return promise;
        }

        function addGroupMembers(groupId, members) {
            var promise = PromiseFactory.createPromise();
            var countToAdd = members.length;
            var added = 0;

            if (countToAdd === 0) {
                promise.resolve(0);
                return promise;
            }

            getGroup(groupId).then(function (group) {
                for (var idx in members) {
                    group.addMember(members[idx]).then(function () {
                        added++;
                        if (added === countToAdd) {
                            promise.resolve(countToAdd);
                        }
                    }, onError.bind(promise));
                }
            }, onError.bind(promise));

            return promise;
        }

        function getApps() {
            var promise = PromiseFactory.createPromise();

            client.applications.getApplications().fetchAll().then(function (apps) {
                promise.resolve(apps);
            }, onError.bind(promise));

            return promise;
        }

        function getApp(objectId) {
            var promise = PromiseFactory.createPromise();

            client.applications.getApplication(objectId).fetch().then(function (app) {
                promise.resolve(app);
            }, onError.bind(promise));

            return promise;
        }

        function createApp(identifierUris, displayName) {
            var app = new AadGraph.Application(AadClient.context, null, null);
            app.displayName = displayName || 'testApp1';
            app.identifierUris = identifierUris;
            return app;
        }

        function addApp(identifierUris, displayName) {
            var promise = PromiseFactory.createPromise();
            var newApp = createApp(identifierUris, displayName);

            client.applications.addApplication(newApp).then(function (app) {
                promise.resolve(app);
            }, onError.bind(promise));

            return promise;
        }

        function editApp(app, identifierUris, displayName) {
            var promise = PromiseFactory.createPromise();

            app.displayName = displayName;
            app.identifierUris = identifierUris;

            app.update().then(function () {
                promise.resolve();
            }, onError.bind(promise));

            return promise;
        }

        function deleteApp(objectId) {
            var promise = PromiseFactory.createPromise();

            getApp(objectId).then(function (app) {
                app.delete().then(function () {
                    promise.resolve();
                }, onError.bind(promise));
            }, onError.bind(promise));

            return promise;
        }

        function getDeletedApps() {
            var promise = PromiseFactory.createPromise();

            client.deletedDirectoryObjects.asApplications().fetchAll().then(function(apps) {
                promise.resolve(apps);
            }, onError.bind(promise));

            return promise;
        }

        function getDeletedApp(objectId) {
            var promise = PromiseFactory.createPromise();

            client.deletedDirectoryObjects.getDirectoryObject(objectId).fetch().then(function (app) {
                promise.resolve(app);
            }, onError.bind(promise));

            return promise;
        }

        function restoreApp(app, identifierUris) {
            var promise = PromiseFactory.createPromise();

            app.restore(identifierUris).then(function(restored) {
                promise.resolve(restored);
            }, onError.bind(promise));

            return promise;
        }

        function authenticate() {
            var promise = PromiseFactory.createPromise();

            logOut(appId).then(function() {
                authContext.acquireTokenAsync(resourceUrl, appId, redirectUrl).then(function(token) {
                    promise.resolve(token);
                },
                onError.bind(promise));
            }, onError.bind(promise));

            return promise;
        }

        function logOut() {
            var promise = PromiseFactory.createPromise();

            authContext.tokenCache.clear().then(function() {
                promise.resolve();
            }, onError.bind(promise));

            return promise;
        }
    }
})();