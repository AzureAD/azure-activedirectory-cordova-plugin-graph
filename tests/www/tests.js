/* global cordova, exports, Exchange, O365Auth, jasmine, describe, it, expect, beforeEach, afterEach, pending */

var TENANT_ID = '17bf7168-5251-44ed-a3cf-37a5997cc451';
var APP_ID = '3cfa20df-bca4-4131-ab92-626fb800ebb5';
var REDIRECT_URL = "http://test.com";

// Used for test entities userPrincipalName generation
var TENANT_NAME = 'testlaboratory.onmicrosoft.com';

var AUTH_URL = 'https://login.windows.net/' + TENANT_ID + '/';
var RESOURCE_URL = 'https://graph.windows.net/';
var ENDPOINT_URL = RESOURCE_URL + TENANT_ID;

var TEST_USER_ID = '';

var AuthenticationContext = Microsoft.ADAL.AuthenticationContext;
var Deferred = require('cordova-plugin-ms-adal.utility').Utility.Deferred;

var guid = function () {
    function _p8(s) {
        var p = (Math.random().toString(16) + "000000000").substr(2, 8);
        return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
    }
    return _p8() + _p8(true) + _p8(true) + _p8();
};

exports.defineAutoTests = function () {
    jasmine.DEFAULT_TIMEOUT_INTERVAL = 30000;

    function createAadGraphClient() {
        return new Microsoft.AADGraph.ActiveDirectoryClient(ENDPOINT_URL,
            new AuthenticationContext(AUTH_URL), RESOURCE_URL, APP_ID, REDIRECT_URL);
    };

    describe('Login: ', function () {
        var authContext, backInterval;
        beforeEach(function () {
            authContext = new AuthenticationContext(AUTH_URL);

            // increase standart jasmine timeout so that user can login
            backInterval = jasmine.DEFAULT_TIMEOUT_INTERVAL;
            jasmine.DEFAULT_TIMEOUT_INTERVAL = 120000;
        });

        afterEach(function () {
            // revert back default jasmine timeout
            jasmine.DEFAULT_TIMEOUT_INTERVAL = backInterval;
        });

        it("login.spec.1 should login first", function (done) {
            authContext.acquireTokenSilentAsync(RESOURCE_URL, APP_ID, TEST_USER_ID).then(function (authResult) {
                console.log("Token is: " + authResult.accessToken);
                expect(authResult).toBeDefined();
                done();
            }, function (err) {
                console.warn("You should login in the manual tests first");

                authContext.acquireTokenAsync(RESOURCE_URL, APP_ID, REDIRECT_URL).then(function (authResult) {
                    console.log("Token is: " + authResult.accessToken);
                    expect(authResult).toBeDefined();
                    done();
                }, function (err) {
                    console.error(err);
                    expect(err).toBeUndefined();
                    done();
                });
            });
        });
    });

    describe('Auth module: ', function () {
        var authContext;

        beforeEach(function () {
            authContext = new AuthenticationContext(AUTH_URL);
        });

        it("auth.spec.1 should contain a Context constructor", function () {
            expect(AuthenticationContext).toBeDefined();
            expect(AuthenticationContext).toEqual(jasmine.any(Function));
        });

        it("auth.spec.2 should successfully create a Context object", function () {
            var fakeAuthUrl = "fakeAuthUrl",
                context = new AuthenticationContext(fakeAuthUrl);

            expect(context).not.toBeNull();
            expect(context).toEqual(jasmine.objectContaining({
                authority: fakeAuthUrl
            }));
        });
    });

    describe('AadGraph client: ', function () {
        var authContext;

        beforeEach(function () {
            authContext = new AuthenticationContext(AUTH_URL);
        });

        it('client.spec.1 should exists', function () {
            expect(Microsoft.AADGraph.ActiveDirectoryClient).toBeDefined();
            expect(Microsoft.AADGraph.ActiveDirectoryClient).toEqual(jasmine.any(Function));
        });

        it('client.spec.2 should be able to create a new client', function () {
            var client = createAadGraphClient();

            expect(client).not.toBe(null);
            expect(client.context).toBeDefined();
            expect(client.context.serviceRootUri).toBeDefined();
            expect(client.context._getAccessTokenFn).toBeDefined();
            expect(client.context.serviceRootUri).toEqual(ENDPOINT_URL);
            expect(client.context._getAccessTokenFn).toEqual(jasmine.any(Function));
        });

        it('client.spec.3 should contain \'directoryObjects\' property', function () {
            var client = createAadGraphClient();

            expect(client.directoryObjects).toBeDefined();
            expect(client.directoryObjects).toEqual(jasmine.any(Microsoft.AADGraph.DirectoryObjects));

            // expect that client.directoryObjects is readonly
            var backupClientDirectoryObjects = client.directoryObjects;
            client.directoryObjects = "somevalue";
            expect(client.directoryObjects).not.toEqual("somevalue");
            expect(client.directoryObjects).toEqual(backupClientDirectoryObjects);
        });
    });

    describe('AAD Graph API: ', function () {
        function fail(done, err) {
            expect(err).toBeUndefined();
            if (err != null) {
                if (err.responseText != null) {
                    expect(err.responseText).toBeUndefined();
                    console.error('Error: ' + err.responseText);
                } else {
                    console.error('Error: ' + err);
                }
            }

            done();
        };

        beforeEach(function () {
            var that = this;
            this.client = createAadGraphClient();

            this.tempEntities = [];

            this.runSafely = function runSafely(testFunc, done) {
                try {
                    // Wrapping the call into try/catch to avoid test suite crashes and `hanging` test entities
                    testFunc(done);
                } catch (err) {
                    fail.call(that, done, err);
                }
            };

            this.createUser = function createUser(displayName) {
                var user = new Microsoft.AADGraph.User();
                displayName = displayName || guid();
                user.displayName = displayName;
                user.accountEnabled = true;
                user.mailNickname = displayName + 'MailNickname';
                user.userPrincipalName = displayName + '@' + TENANT_NAME;
                var passwordProfile = new Microsoft.AADGraph.PasswordProfile();
                passwordProfile.password = "Test1234";
                passwordProfile.forceChangePasswordNextLogin = false;
                user.passwordProfile = passwordProfile;
                user.usageLocation = 'US'; // This property is needed for `assignLicense` tests
                return user;
            };

            this.createGroup = function createGroup(displayName) {
                var group = new Microsoft.AADGraph.Group();
                group.displayName = displayName || 'testGroup1';
                group.mailNickname = 'test';
                group.mailEnabled = 'false';
                group.securityEnabled = 'true';
                return group;
            };

            this.createApp = function createApp(displayName) {
                var app = new Microsoft.AADGraph.Application();
                displayName = displayName || 'AppToTest';
                var identifierUrl = 'https://localhost:5362/' + displayName;
                app.displayName = displayName;
                app.identifierUris = [identifierUrl];
                return app;
            };

            this.createServicePrincipal = function createServicePrincipal(app) {
                var principal = new Microsoft.AADGraph.ServicePrincipal();
                principal.appId = app.appId;
                principal.servicePrincipalNames = app.identifierUris;
                return principal;
            };

            this.createGrant = function createGrant(resourceId, clientId, principalId) {
                var grant = new Microsoft.AADGraph.OAuth2PermissionGrant();
                grant.resourceId = resourceId;
                grant.clientId = clientId;
                grant.principalId = principalId;
                grant.consentType = 'Principal';
                grant.startTime = '2014-03-01';
                grant.expiryTime = '2014-04-01';
                return grant;
            };

            this.createDevice = function createDevice() {
                var device = new Microsoft.AADGraph.Device();
                device.displayName = guid();
                device.deviceId = guid();
                device.accountEnabled = true;

                var altSecId = new Microsoft.AADGraph.AlternativeSecurityId();
                altSecId.key = btoa(guid());
                altSecId.type = 2;
                altSecId.identityProvider = null;

                device.alternativeSecurityIds = [altSecId];
                device.deviceOSType = 'Windows Phone';
                device.deviceOSVersion = '8.1';
                return device;
            };

            this.createAppRoleAssignment = function createAppRoleAssignment(servicePrincipal, user) {
                var assignment = new Microsoft.AADGraph.AppRoleAssignment();
                assignment.id = '00000000-0000-0000-0000-000000000000';//default guid
                assignment.resourceId = servicePrincipal.objectId;
                assignment.principalId = user.objectId;
                return assignment;
            };

            this.createExtensionProperty = function createExtensionProperty(name) {
                var property = new Microsoft.AADGraph.ExtensionProperty();
                property.name = name;
                property.dataType = 'String';
                property.targetObjects = property.targetObjects.concat('User');
                return property;
            };

            this.addANumberOfUsers = function addANumberOfUsers(count, namePrefix) {
                var tempArray = [];
                namePrefix = namePrefix || 'test';
                var deferred = new Deferred();
                var resolve = function (obj) {
                    deferred.resolve(obj);
                };
                var reject = function (err) {
                    deferred.reject(err);
                };

                var k = 0;
                var errorOccured = false;
                for (var i = 0; i < count; i++) {
                    if (errorOccured) {
                        break;
                    }

                    that.client.users.addUser(this.createUser(namePrefix + guid())).then(function (user) {
                        tempArray.push(user);
                        k++;
                        if (k === count) {
                            resolve(tempArray);
                        }
                    }, function (ex) {
                        i = count + 1; // End the loop

                        if (!errorOccured) {
                            errorOccured = true;

                            reject(ex);
                        }
                    });
                }

                return deferred;
            };

            jasmine.Expectation.addMatchers({
                toContainObjWithId: function () {
                    return {
                        compare: function (arr, objectId) {
                            return {
                                pass: arr.filter(function (obj) {
                                    return obj.objectId === objectId;
                                }).length > 0
                            };
                        }
                    };
                }
            });
        });

        afterEach(function (done) {
            var removedEntitiesCount = 0;
            var entitiesToRemoveCount = this.tempEntities.length;

            if (entitiesToRemoveCount === 0) {
                done();
            } else {
                this.tempEntities.forEach(function (entity) {
                    try {
                        entity.delete().then(function () {
                            removedEntitiesCount++;
                            if (removedEntitiesCount === entitiesToRemoveCount) {
                                done();
                            }
                        }, function (err) {
                            expect(err).toBeUndefined();
                            done();
                        });
                    } catch (e) {
                        expect(e).toBeUndefined();
                        done();
                    }
                });
            }
        });

        describe('Groups', function () {
            it("groups.spec.1 should be able to create a new group", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (added) {
                        that.tempEntities.push(added);
                        expect(added.objectId).toBeDefined();
                        expect(added.path).toMatch(added.objectId);
                        expect(added).toEqual(jasmine.any(Microsoft.AADGraph.Group));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.2 should be able to get groups", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.getGroups().fetchAll().then(function (groups) {
                        expect(groups).toBeDefined();
                        expect(groups).toEqual(jasmine.any(Array));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.2.1 should be able to get groups (tries to add a group first)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.groups.getGroups().fetchAll().then(function (groups) {
                            expect(groups).toBeDefined();
                            expect(groups).toEqual(jasmine.any(Array));
                            expect(groups.length).toBeGreaterThan(0);
                            expect(groups[0]).toEqual(jasmine.any(Microsoft.AADGraph.Group));
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.3 should be able to get group by Id", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.groups.getGroups().fetchAll().then(function (groups) {
                            var addedGroup = groups[0];
                            that.client.groups.getGroup(addedGroup._objectId).fetch().then(function (groupFoundById) {
                                expect(groupFoundById.objectId).toBe(addedGroup._objectId);
                                expect(groupFoundById.path).toMatch(addedGroup.objectId);
                                expect(groupFoundById).toEqual(jasmine.any(Microsoft.AADGraph.Group));
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.4 should be able to apply filter to groups", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (created) {
                        that.tempEntities.push(created);
                        var filter = 'displayName eq \'' + created.displayName + '\'';
                        that.client.groups.getGroups().filter(filter).fetchAll().then(function (groups) {
                            expect(groups).toBeDefined();
                            expect(groups).toEqual(jasmine.any(Array));
                            expect(groups.length).toBeGreaterThan(0);
                            expect(groups[0]).toEqual(jasmine.any(Microsoft.AADGraph.Group));
                            expect(groups[0].displayName).toEqual(created.displayName);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.5 should be able to get a newly created group by objectId", function (done) {
                var that = this;

                that.runSafely(function () {
                    var group = that.createGroup();
                    that.client.groups.addGroup(group).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.groups.getGroup(added.objectId).fetch().then(function (got) {
                            expect(got.userPrincipalName).toEqual(group.userPrincipalName);
                            expect(got.displayName).toEqual(group.displayName);
                            expect(got.mailNickname).toEqual(group.mailNickname);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.6 should be able to modify an existing group", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (added) {
                        that.tempEntities.push(added);
                        added.displayName = guid();
                        added.update().then(function () {
                            that.client.groups.getGroup(added.objectId).fetch().then(function (got) {
                                expect(got.objectId).toEqual(added.objectId);
                                expect(got.displayName).toEqual(added.displayName);
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.7 should be able to delete an existing group", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (added) {
                        added.delete().then(function () {
                            that.client.groups.getGroup(added.objectId).fetch().then(function (got) {
                                expect(got).toBeUndefined();
                                done();
                            }, function (err) {
                                expect(err.statusText).toBeDefined();
                                expect(err.statusText).toMatch("Not Found");
                                done();
                            });
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.8 should be able to add a group member", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup('testGroup1')).then(function (rootGroup) {
                        that.tempEntities.push(rootGroup);
                        that.client.groups.addGroup(that.createGroup('testGroup2')).then(function (nestedGroup) {
                            that.tempEntities.push(nestedGroup);
                            rootGroup.addMember(nestedGroup).then(function () {
                                rootGroup.members.getDirectoryObjects().fetchAll().then(function (members) {
                                    expect(members).toEqual(jasmine.any(Array));
                                    expect(members.length).toEqual(1);
                                    expect(members[0]).toEqual(jasmine.any(Microsoft.AADGraph.Group));
                                    expect(members[0].objectId).toEqual(nestedGroup.objectId);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.9 should be able to add a nested user member", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup('testGroup1')).then(function (createdGroup1) {
                        that.tempEntities.push(createdGroup1);
                        that.client.groups.addGroup(that.createGroup('testGroup2')).then(function (createdGroup2) {
                            that.tempEntities.push(createdGroup2);
                            that.client.users.addUser(that.createUser()).then(function (createdUser) {
                                that.tempEntities.push(createdUser);

                                createdGroup1.addMember(createdGroup2).then(function () {
                                    createdGroup1.addMember(createdUser).then(function () {
                                        createdGroup1.members.getDirectoryObjects().fetchAll().then(function (members) {
                                            expect(members).toEqual(jasmine.any(Array));
                                            expect(members.length).toEqual(2);
                                            expect(members).toContain(jasmine.any(Microsoft.AADGraph.Group));
                                            expect(members).toContainObjWithId(createdGroup2.objectId);
                                            expect(members).toContain(jasmine.any(Microsoft.AADGraph.User));
                                            expect(members).toContainObjWithId(createdUser.objectId);

                                            done();
                                        });
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.10 should be able to add nested groups", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup('testGroup1')).then(function (createdGroup1) {
                        that.tempEntities.push(createdGroup1);
                        that.client.groups.addGroup(that.createGroup('testGroup2')).then(function (createdGroup2) {
                            that.tempEntities.push(createdGroup2);
                            that.client.groups.addGroup(that.createGroup('testGroup3')).then(function (createdGroup3) {
                                that.tempEntities.push(createdGroup3);

                                createdGroup2.addMember(createdGroup3).then(function () {
                                    createdGroup1.addMember(createdGroup2).then(function () {
                                        createdGroup1.members.getDirectoryObjects().fetchAll().then(function (members) {
                                            expect(members).toEqual(jasmine.any(Array));
                                            expect(members.length).toEqual(1);
                                            expect(members).toContain(jasmine.any(Microsoft.AADGraph.Group));
                                            expect(members).toContainObjWithId(createdGroup2.objectId);

                                            createdGroup2.members.getDirectoryObjects().fetchAll().then(function (nestedMembers) {
                                                expect(nestedMembers).toEqual(jasmine.any(Array));
                                                expect(nestedMembers.length).toEqual(1);
                                                expect(nestedMembers).toContain(jasmine.any(Microsoft.AADGraph.Group));
                                                expect(nestedMembers).toContainObjWithId(createdGroup3.objectId);

                                                done();
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("groups.spec.11 should be able to remove a group member", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup('testGroup1')).then(function (rootGroup) {
                        that.tempEntities.push(rootGroup);
                        that.client.groups.addGroup(that.createGroup('testGroup2')).then(function (nestedGroup) {
                            that.tempEntities.push(nestedGroup);
                            rootGroup.addMember(nestedGroup).then(function () {
                                rootGroup.members.getDirectoryObjects().fetchAll().then(function (members) {
                                    expect(members).toEqual(jasmine.any(Array));
                                    expect(members.length).toEqual(1);
                                    expect(members[0]).toEqual(jasmine.any(Microsoft.AADGraph.Group));
                                    expect(members[0].objectId).toEqual(nestedGroup.objectId);

                                    rootGroup.deleteMember(nestedGroup).then(function () {
                                        rootGroup.members.getDirectoryObjects().fetchAll().then(function (membersAfterDeletion) {
                                            expect(membersAfterDeletion).toEqual(jasmine.any(Array));
                                            expect(membersAfterDeletion.length).toEqual(0);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Users", function () {
            it("users.spec.1 should be able to create a new user", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (added) {
                        that.tempEntities.push(added);
                        expect(added.objectId).toBeDefined();
                        expect(added.path).toMatch(added.objectId);
                        expect(added).toEqual(jasmine.any(Microsoft.AADGraph.User));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.2 should be able to get users", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.getUsers().fetchAll().then(function (users) {
                        expect(users).toBeDefined();
                        expect(users).toEqual(jasmine.any(Array));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.2.1 should be able to get users (tries to add a user first)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.users.getUsers().fetchAll().then(function (users) {
                            expect(users).toBeDefined();
                            expect(users).toEqual(jasmine.any(Array));
                            expect(users.length).toBeGreaterThan(0);
                            expect(users[0]).toEqual(jasmine.any(Microsoft.AADGraph.User));
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.3 should be able to apply filter to users", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.users.getUsers().filter("displayName eq '" + created.displayName + "'").fetchAll().then(function (users) {
                            expect(users).toBeDefined();
                            expect(users).toEqual(jasmine.any(Array));
                            expect(users.length).toEqual(1);
                            expect(users[0]).toEqual(jasmine.any(Microsoft.AADGraph.User));
                            expect(users[0].displayName).toEqual(created.displayName);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.4 should be able to apply top query to users", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.users.addUser(that.createUser()).then(function (created2) {
                            that.tempEntities.push(created2);
                            //You should not use `fetchAll()` when you apply `top`.
                            that.client.users.getUsers().top(1).fetch().then(function (users) {
                                expect(users).toBeDefined();
                                expect(users.currentPage).toBeDefined();
                                expect(users.currentPage).toEqual(jasmine.any(Array));
                                expect(users.currentPage.length).toEqual(1);
                                expect(users.currentPage[0]).toEqual(jasmine.any(Microsoft.AADGraph.User));
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.5 should be able to get a newly created user by objectId", function (done) {
                var that = this;
                var newUser = that.createUser();

                that.runSafely(function () {
                    that.client.users.addUser(newUser).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.users.getUser(added.objectId).fetch().then(function (got) {
                            expect(got.objectId).toEqual(added.objectId);
                            expect(got).toEqual(jasmine.any(Microsoft.AADGraph.User));
                            expect(got.userPrincipalName).toEqual(newUser.userPrincipalName);
                            expect(got.displayName).toEqual(newUser.displayName);
                            expect(got.mailNickname).toEqual(newUser.mailNickname);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.6 should be able to modify existing user", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (added) {
                        that.tempEntities.push(added);
                        added.displayName = guid();
                        added.update().then(function () {
                            that.client.users.getUser(added.objectId).fetch().then(function (got) {
                                expect(got.objectId).toEqual(added.objectId);
                                expect(got.displayName).toEqual(added.displayName);
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.7 should be able to delete existing user", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (added) {
                        added.delete().then(function () {
                            that.client.users.getUser(added.objectId).fetch().then(function (got) {
                                expect(got).toBeUndefined();
                                done();
                            }, function (err) {
                                expect(err.statusText).toBeDefined();
                                expect(err.statusText).toMatch("Not Found");
                                done();
                            });
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.8 user should be able to update manager", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (user1) {
                        that.tempEntities.push(user1);
                        that.client.users.addUser(that.createUser()).then(function (user2) {
                            that.tempEntities.push(user2);
                            that.client.users.addUser(that.createUser()).then(function (user3) {
                                that.tempEntities.push(user3);
                                user1.update_manager(user2).then(function () {
                                    user1.manager.fetch().then(function (user1Manager) {
                                        expect(user1Manager).toBeDefined();
                                        expect(user1Manager).toEqual(jasmine.any(Microsoft.AADGraph.User));
                                        expect(user1Manager.objectId).toEqual(user2.objectId);

                                        user1.update_manager(user3).then(function () {
                                            user1.manager.fetch().then(function (user1UpdatedManager) {
                                                expect(user1UpdatedManager).toBeDefined();
                                                expect(user1UpdatedManager).toEqual(jasmine.any(Microsoft.AADGraph.User));
                                                expect(user1UpdatedManager.objectId).toEqual(user3.objectId);

                                                done();
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.9 user should be able to get manager", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (user1) {
                        that.tempEntities.push(user1);
                        that.client.users.addUser(that.createUser()).then(function (user2) {
                            that.tempEntities.push(user2);
                            user1.update_manager(user2).then(function () {
                                user1.manager.fetch().then(function (user1Manager) {
                                    expect(user1Manager).toBeDefined();
                                    expect(user1Manager).toEqual(jasmine.any(Microsoft.AADGraph.User));
                                    expect(user1Manager.objectId).toEqual(user2.objectId);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.10 user should be able to get direct reports", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (user1) {
                        that.tempEntities.push(user1);
                        that.client.users.addUser(that.createUser()).then(function (user2) {
                            that.tempEntities.push(user2);
                            user1.update_manager(user2).then(function () {
                                user2.directReports.getDirectoryObjects().fetchAll().then(function (user2DirectReports) {
                                    expect(user2DirectReports).toBeDefined();
                                    expect(user2DirectReports).toEqual(jasmine.any(Array));
                                    expect(user2DirectReports).toContainObjWithId(user1.objectId);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("users.spec.11 user should be able to reset password", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser()).then(function (user1) {
                        that.tempEntities.push(user1);
                        user1.passwordProfile = new Microsoft.AADGraph.PasswordProfile();
                        user1.passwordProfile.password = "ChangedPass1234";
                        user1.passwordProfile.forceChangePasswordNextLogin = false;
                        user1.update().then(function () {
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Roles", function () {
            it("roles.spec.1 should be able to get roles", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.directoryRoles.getDirectoryRoles().fetchAll().then(function (roles) {
                        expect(roles).toBeDefined();
                        expect(roles).toEqual(jasmine.any(Array));
                        expect(roles[0]).toEqual(jasmine.any(Microsoft.AADGraph.DirectoryRole));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("roles.spec.2 should be able to get role by objectId", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.directoryRoles.getDirectoryRoles().fetchAll().then(function (roles) {
                        var role = roles[0];
                        that.client.directoryRoles.getDirectoryRole(role.objectId).fetch().then(function (got) {
                            expect(got.objectType).toEqual(role.objectType);
                            expect(got.objectId).toEqual(role.objectId);
                            expect(got.description).toEqual(role.description);
                            expect(got).toEqual(jasmine.any(Microsoft.AADGraph.DirectoryRole));
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("roles.spec.3 should be able to add user to role", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.directoryRoles.getDirectoryRoles().fetchAll().then(function (roles) {
                        var role = roles[0];

                        that.client.users.addUser(that.createUser()).then(function (user) {
                            that.tempEntities.push(user);
                            role.addMember(user).then(function () {
                                role.members.getDirectoryObjects().fetchAll().then(function (members) {
                                    expect(members).toContainObjWithId(user.objectId);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("roles.spec.4 should be able to delete user from role", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.directoryRoles.getDirectoryRoles().fetchAll().then(function (roles) {
                        var role = roles[0];

                        that.client.users.addUser(that.createUser()).then(function (user) {
                            that.tempEntities.push(user);
                            role.members.getDirectoryObjects().fetchAll().then(
                                function (membersBefore) {
                                    role.addMember(user).then(function () {
                                        role.deleteMember(user).then(function () {
                                            role.members.getDirectoryObjects().fetchAll().then(function (membersAfter) {
                                                expect(membersBefore.length).toBe(membersAfter.length);
                                                expect(membersAfter).not.toContainObjWithId(user.objectId);
                                                done();
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Additional functions", function () {
            it("additional funcs.spec.1 should be able to execute isMemberOf (user)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group) {
                        that.tempEntities.push(group);
                        that.client.users.addUser(that.createUser()).then(function (user) {
                            that.tempEntities.push(user);
                            group.addMember(user).then(function () {
                                that.client.isMemberOf(group.objectId, user.objectId).then(function (obj) {
                                    expect(obj).toBeDefined();
                                    expect(obj.value).toBe(true);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.1.1 should be able to execute isMemberOf (group)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            group1.addMember(group2).then(function () {
                                that.client.isMemberOf(group1.objectId, group2.objectId).then(function (obj) {
                                    expect(obj).toBeDefined();
                                    expect(obj.value).toBe(true);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.1.2 group isMemberOf should be transitive", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.groups.addGroup(that.createGroup()).then(function (group3) {
                                that.tempEntities.push(group3);
                                group1.addMember(group2).then(function () {
                                    group2.addMember(group3).then(function () {
                                        that.client.isMemberOf(group1.objectId, group3.objectId).then(function (obj) {
                                            expect(obj).toBeDefined();
                                            expect(obj.value).toBe(true);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.2 should be able to execute getMemberGroups (user)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group) {
                        that.tempEntities.push(group);
                        that.client.users.addUser(that.createUser()).then(function (user) {
                            that.tempEntities.push(user);
                            group.addMember(user).then(function () {
                                user.getMemberGroups(false).then(function (groupIds) {
                                    expect(groupIds).toBeDefined();
                                    expect(groupIds).toEqual(jasmine.any(Array));
                                    expect(groupIds.length).toBe(1);
                                    expect(groupIds).toContain(group.objectId);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.2.1 should be able to execute getMemberGroups (group)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            group1.addMember(group2).then(function () {
                                group2.getMemberGroups(false).then(function (groupIds) {
                                    expect(groupIds).toBeDefined();
                                    expect(groupIds).toEqual(jasmine.any(Array));
                                    expect(groupIds.length).toBe(1);
                                    expect(groupIds).toContain(group1.objectId);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.2.2 group getMemberGroups should be transitive", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.groups.addGroup(that.createGroup()).then(function (group3) {
                                that.tempEntities.push(group3);
                                group1.addMember(group2).then(function () {
                                    group2.addMember(group3).then(function () {
                                        group3.getMemberGroups(false).then(function (groupIds) {
                                            expect(groupIds).toBeDefined();
                                            expect(groupIds).toEqual(jasmine.any(Array));
                                            expect(groupIds.length).toBe(2);
                                            expect(groupIds).toContain(group1.objectId);
                                            expect(groupIds).toContain(group2.objectId);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.3 getMemberGroups should work for a user added to several groups", function (done) {
                var that = this;
                var group1, group2;

                that.runSafely(function () {
                    group1 = that.createGroup();
                    that.client.groups.addGroup(group1).then(function (group1Entity) {
                        that.tempEntities.push(group1Entity);
                        that.client.users.addUser(that.createUser()).then(function (user) {
                            that.tempEntities.push(user);
                            group2 = that.createGroup();
                            that.client.groups.addGroup(group2).then(function (group2Entity) {
                                that.tempEntities.push(group2Entity);
                                group1Entity.addMember(user).then(function () {
                                    group2Entity.addMember(user).then(function () {
                                        user.getMemberGroups(false).then(function (groupIds) {
                                            expect(groupIds).toBeDefined();
                                            expect(groupIds).toEqual(jasmine.any(Array));
                                            expect(groupIds.length).toEqual(2);
                                            expect(groupIds).toContain(group1Entity.objectId);
                                            expect(groupIds).toContain(group2Entity.objectId);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.4 should be able to execute getMemberObjects", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group) {
                        that.tempEntities.push(group);
                        that.client.users.addUser(that.createUser()).then(function (user) {
                            that.tempEntities.push(user);
                            group.addMember(user).then(function () {
                                user.getMemberObjects(false).then(function (objectIds) {
                                    expect(objectIds).toBeDefined();
                                    expect(objectIds).toEqual(jasmine.any(Array));
                                    expect(objectIds.length).toBe(1);
                                    expect(objectIds).toContain(group.objectId);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.4.1 getMemberObjects should support roles", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.directoryRoles.getDirectoryRoles().fetchAll().then(function (roles) {
                        var role = roles[0];
                        that.client.groups.addGroup(that.createGroup()).then(function (group) {
                            that.tempEntities.push(group);
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                group.addMember(user).then(function () {
                                    role.addMember(user).then(function () {
                                        user.getMemberObjects(false).then(function (objectIds) {
                                            expect(objectIds).toBeDefined();
                                            expect(objectIds).toEqual(jasmine.any(Array));
                                            expect(objectIds.length).toBe(2);
                                            expect(objectIds).toContain(group.objectId);
                                            expect(objectIds).toContain(role.objectId);
                                            role.deleteMember(user).then(done, function (err) {
                                                fail.call(that, done, 'additional funcs.spec.4.1 The member was not deleted from role because of an error: ' + err);
                                            });
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.4.2 getMemberObjects should be transitive", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                group1.addMember(group2).then(function () {
                                    group2.addMember(user).then(function () {
                                        user.getMemberObjects(false).then(function (objectIds) {
                                            expect(objectIds).toBeDefined();
                                            expect(objectIds).toEqual(jasmine.any(Array));
                                            expect(objectIds.length).toBe(2);
                                            expect(objectIds).toContain(group1.objectId);
                                            expect(objectIds).toContain(group2.objectId);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.4.3 getMemberObjects should return only correct objects", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                group1.addMember(user).then(function () {
                                    user.getMemberObjects(false).then(function (objectIds) {
                                        expect(objectIds).toBeDefined();
                                        expect(objectIds).toEqual(jasmine.any(Array));
                                        expect(objectIds.length).toBe(1);
                                        expect(objectIds).toContain(group1.objectId);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.5 should be able to execute checkMemberGroups (single entry)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                group1.addMember(user).then(function () {
                                    user.checkMemberGroups([group1.objectId, group2.objectId]).then(function (groupIds) {
                                        expect(groupIds).toBeDefined();
                                        expect(groupIds).toEqual(jasmine.any(Array));
                                        expect(groupIds.length).toBe(1);
                                        expect(groupIds).toContain(group1.objectId);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.5.1 group checkMemberGroups should be transitive", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.groups.addGroup(that.createGroup()).then(function (group3) {
                                that.tempEntities.push(group3);
                                group1.addMember(group2).then(function () {
                                    group2.addMember(group3).then(function () {
                                        group3.checkMemberGroups([group1.objectId, group2.objectId]).then(function (groupIds) {
                                            expect(groupIds).toBeDefined();
                                            expect(groupIds).toEqual(jasmine.any(Array));
                                            expect(groupIds.length).toBe(2);
                                            expect(groupIds).toContain(group1.objectId);
                                            expect(groupIds).toContain(group2.objectId);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.6 should be able to execute checkMemberGroups (multiple entries)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                group1.addMember(user).then(function () {
                                    group2.addMember(user).then(function () {
                                        user.checkMemberGroups([group1.objectId, group2.objectId]).then(function (groupIds) {
                                            expect(groupIds).toBeDefined();
                                            expect(groupIds).toEqual(jasmine.any(Array));
                                            expect(groupIds.length).toBe(2);
                                            expect(groupIds).toContain(group1.objectId);
                                            expect(groupIds).toContain(group2.objectId);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.7 should be able to execute user' memberOf", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                group1.addMember(user).then(function () {
                                    user.memberOf.getDirectoryObjects().fetchAll().then(function (parentGroups) {
                                        expect(parentGroups).toBeDefined();
                                        expect(parentGroups).toEqual(jasmine.any(Array));
                                        expect(parentGroups.length).toBe(1);
                                        expect(parentGroups).toContainObjWithId(group1.objectId);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.8 user' memberOf should be intransitive", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.groups.addGroup(that.createGroup()).then(function (group1) {
                        that.tempEntities.push(group1);
                        that.client.groups.addGroup(that.createGroup()).then(function (group2) {
                            that.tempEntities.push(group2);
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                group1.addMember(group2).then(function () {
                                    group2.addMember(user).then(function () {
                                        user.memberOf.getDirectoryObjects().fetchAll().then(function (parentGroups) {
                                            expect(parentGroups).toBeDefined();
                                            expect(parentGroups).toEqual(jasmine.any(Array));
                                            expect(parentGroups.length).toBe(1);
                                            expect(parentGroups).toContainObjWithId(group2.objectId);
                                            expect(parentGroups).not.toContainObjWithId(group1.objectId);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.9 assignLicense should be able assign license", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.subscribedSkus.getSubscribedSkus().fetchAll().then(function (subscribedSkus) {
                        if (subscribedSkus.length === 0 || subscribedSkus[0].servicePlans._array.length === 0) {
                            fail.call(that, done, 'Please add subscriptions and service plans to run this test');
                        } else {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                var servicePlanId = subscribedSkus[0].servicePlans._array[0].servicePlanId;
                                user.assignLicense([
                                    {
                                        "disabledPlans": [servicePlanId],
                                        "skuId": subscribedSkus[0].skuId
                                    }
                                ], []).then(function () {
                                    that.client.users.getUser(user.objectId).fetch().then(function (updatedUser) {
                                        var licenses = updatedUser.assignedLicenses._array;
                                        expect(licenses.length).toBe(1);
                                        expect(licenses[0].skuId).toBe(subscribedSkus[0].skuId);
                                        expect(licenses[0].disabledPlans.length).toBe(1);
                                        expect(licenses[0].disabledPlans[0]).toBe(servicePlanId);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }
                    }, fail.bind(that, done));
                }, done);
            });

            it("additional funcs.spec.9.1 assignLicense should be able remove license", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.subscribedSkus.getSubscribedSkus().fetchAll().then(function (subscribedSkus) {
                        if (subscribedSkus.length === 0 || subscribedSkus[0].servicePlans._array.length === 0) {
                            fail.call(that, done, 'Please add subscriptions and service plans to run this test');
                        } else {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                var servicePlanId = subscribedSkus[0].servicePlans._array[0].servicePlanId;
                                user.assignLicense([
                                    {
                                        "disabledPlans": [servicePlanId],
                                        "skuId": subscribedSkus[0].skuId
                                    }
                                ], []).then(function () {
                                    user.assignLicense([], [subscribedSkus[0].skuId]).then(function () {
                                        that.client.users.getUser(user.objectId).fetch().then(function (updatedUser) {
                                            expect(updatedUser.assignedLicenses.length).toBe(0);
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("SubscribedSku", function () {
            it("subscribedSku.spec.1 should be able to get SubscribedSkus", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.subscribedSkus.getSubscribedSkus().fetchAll().then(function (skus) {
                        expect(skus).toBeDefined();
                        expect(skus).toEqual(jasmine.any(Array));
                        if (skus.length === 0) {
                            fail.call(that, done, 'Please add subscriptions to test this function');
                        } else {
                            expect(skus[0]).toEqual(jasmine.any(Microsoft.AADGraph.SubscribedSku));
                            done();
                        }
                    }, fail.bind(that, done));
                }, done);
            });

            it("subscribedSku.spec.2 should be able to get SubscribedSku by objectId", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.subscribedSkus.getSubscribedSkus().fetchAll().then(function (skus) {
                        expect(skus).toBeDefined();
                        if (skus.length === 0) {
                            fail.call(that, done, 'Please add subscriptions to test this function');
                        } else {
                            that.client.subscribedSkus.getSubscribedSku(skus[0].objectId).fetch().then(function (sku) {
                                expect(sku.objectId).toBe(skus[0].objectId);
                                expect(sku).toEqual(jasmine.any(Microsoft.AADGraph.SubscribedSku));
                                expect(sku.skuId).toBeDefined();
                            }, fail.bind(that, done));
                            done();
                        }
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("TenantDetail", function () {
            it("tenantDetail.spec.1 should be able to get TenantDetails", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.tenantDetails.getTenantDetails().fetchAll().then(function (details) {
                        expect(details).toBeDefined();
                        expect(details).toEqual(jasmine.any(Array));
                        expect(details[0]).toEqual(jasmine.any(Microsoft.AADGraph.TenantDetail));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("tenantDetail.spec.2 should be able to get TenantDetail by objectId", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.tenantDetails.getTenantDetails().fetchAll().then(function (details) {
                        that.client.tenantDetails.getTenantDetail(details[0].objectId).fetch().then(function (detail) {
                            expect(detail.objectId).toBe(details[0].objectId);
                            expect(detail).toEqual(jasmine.any(Microsoft.AADGraph.TenantDetail));
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("tenantDetail.spec.3 should be able to update TenantDetail properties", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.tenantDetails.getTenantDetails().fetchAll().then(function (details) {
                        var detail = details[0];
                        var oldTechMails = detail.technicalNotificationMails;
                        var oldMarketMails = detail.marketingNotificationEmails;

                        var marketEmail = 'test-3-1@microsoft.com';
                        var techEmail = 'test-3-2@microsoft.com';

                        detail.marketingNotificationEmails = detail.marketingNotificationEmails.concat([marketEmail]);
                        detail.technicalNotificationMails = detail.technicalNotificationMails.concat([techEmail]);
                        detail.update().then(function () {
                            that.client.tenantDetails.getTenantDetail(detail.objectId).fetch().then(function (updatedDetail) {
                                expect(updatedDetail.marketingNotificationEmails.length).toBe(oldMarketMails.length + 1);
                                expect(updatedDetail.technicalNotificationMails.length).toBe(oldTechMails.length + 1);
                                expect(updatedDetail.marketingNotificationEmails).toContain(marketEmail);
                                expect(updatedDetail.technicalNotificationMails).toContain(techEmail);
                                detail.marketingNotificationEmails = oldMarketMails;
                                detail.technicalNotificationMails = oldTechMails;
                                detail.update().then(done, function (err) {
                                    fail.call(that, done, 'tenantDetail.spec.3 could not reset tenant\' marketingNotificationEmails and technicalNotificationMails to original values because of an error: ' + err);
                                });
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Applications", function () {
            it("apps.spec.1 should be able to create a new app", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (added) {
                        that.tempEntities.push(added);
                        expect(added.objectId).toBeDefined();
                        expect(added.path).toMatch(added.objectId);
                        expect(added).toEqual(jasmine.any(Microsoft.AADGraph.Application));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.2 should be able to get apps", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.getApplications().fetchAll().then(function (apps) {
                        expect(apps).toBeDefined();
                        expect(apps).toEqual(jasmine.any(Array));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.2.1 should be able to get apps (tries to add a app first)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.applications.getApplications().fetchAll().then(function (applications) {
                            expect(applications).toBeDefined();
                            expect(applications).toEqual(jasmine.any(Array));
                            expect(applications.length).toBeGreaterThan(0);
                            expect(applications[0]).toEqual(jasmine.any(Microsoft.AADGraph.Application));
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.3 should be able to apply filter to apps", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.applications.getApplications().filter("displayName eq '" + created.displayName + "'").fetchAll().then(function (applications) {
                            expect(applications).toBeDefined();
                            expect(applications).toEqual(jasmine.any(Array));
                            expect(applications.length).toEqual(1);
                            expect(applications[0]).toEqual(jasmine.any(Microsoft.AADGraph.Application));
                            expect(applications[0].displayName).toEqual(created.displayName);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.4 should be able to apply top query to apps", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (created) {
                        that.tempEntities.push(created);
                        var secondApp = that.createApp('app2');
                        that.client.applications.addApplication(secondApp).then(function (created2) {
                            that.tempEntities.push(created2);
                            //You should not use `fetchAll()` when you apply `top` to apps.
                            that.client.applications.getApplications().top(1).fetch().then(function (applications) {
                                expect(applications).toBeDefined();
                                expect(applications.currentPage).toBeDefined();
                                expect(applications.currentPage).toEqual(jasmine.any(Array));
                                expect(applications.currentPage.length).toEqual(1);
                                expect(applications.currentPage[0]).toEqual(jasmine.any(Microsoft.AADGraph.Application));
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.5 should be able to get a newly created app by objectId", function (done) {
                var that = this;
                var newApp = that.createApp();

                that.runSafely(function () {
                    that.client.applications.addApplication(newApp).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.applications.getApplication(added.objectId).fetch().then(function (got) {
                            expect(got.objectId).toEqual(added.objectId);
                            expect(got).toEqual(jasmine.any(Microsoft.AADGraph.Application));
                            expect(got.displayName).toEqual(added.displayName);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.6 should be able to modify existing app", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (added) {
                        that.tempEntities.push(added);
                        added.displayName = guid();
                        added.update().then(function () {
                            that.client.applications.getApplication(added.objectId).fetch().then(function (got) {
                                expect(got.objectId).toEqual(added.objectId);
                                expect(got.displayName).toEqual(added.displayName);
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.7 should be able to delete existing app", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (added) {
                        added.delete().then(function () {
                            that.client.applications.getApplication(added.objectId).fetch().then(function (got) {
                                expect(got).toBeUndefined();
                                done();
                            }, function (err) {
                                expect(err.statusText).toBeDefined();
                                expect(err.statusText).toMatch("Not Found");
                                done();
                            });
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.8 should be able to restore deleted app with ommited identifierUris", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (added) {
                        that.tempEntities.push(added);
                        added.delete().then(function () {
                            that.client.deletedDirectoryObjects.asApplications().filter("appId eq '" + added.appId + "'").fetchAll().then(function (apps) {
                                apps[0].restore().then(function (restoredApp) {
                                    expect(restoredApp.objectId).toBe(added.objectId);
                                    that.client.applications.getApplication(added.objectId).fetch().then(function (app) {
                                        expect(app.objectId).toBe(added.objectId);
                                        expect(app.appId).toBe(added.appId);
                                        expect(app.identifierUris.length).toBe(1);
                                        expect(app.identifierUris).toContain(added.identifierUris[0]);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("apps.spec.9 should be able to restore deleted app using identifierUris argument", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (added) {
                        that.tempEntities.push(added);
                        added.delete().then(function () {
                            that.client.deletedDirectoryObjects.asApplications().filter("appId eq '" + added.appId + "'").fetchAll().then(function (apps) {
                                var newUrl = 'https://localhost:5362/restoredtestapp';
                                apps[0].restore([newUrl]).then(function (restoredApp) {
                                    expect(restoredApp.objectId).toBe(added.objectId);
                                    expect(restoredApp.identifierUris.length).toBe(1);
                                    expect(restoredApp.identifierUris).not.toContain(added.identifierUris[0]);
                                    expect(restoredApp.identifierUris).toContain(newUrl);
                                    that.client.applications.getApplication(added.objectId).fetch().then(function (app) {
                                        expect(app.objectId).toBe(added.objectId);
                                        expect(app.appId).toBe(added.appId);
                                        expect(app.identifierUris.length).toBe(1);
                                        expect(app.identifierUris).not.toContain(added.identifierUris[0]);
                                        expect(app.identifierUris).toContain(newUrl);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("ServicePrincipal", function () {
            it("servicePrincipal.spec.1 should be able to create a new service principal", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals')).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (added) {
                            expect(added.objectId).toBeDefined();
                            expect(added.appId).toMatch(app.appId);
                            expect(added.path).toMatch(added.objectId);
                            expect(added).toEqual(jasmine.any(Microsoft.AADGraph.ServicePrincipal));
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("servicePrincipal.spec.2 should be able to get service principals", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.servicePrincipals.getServicePrincipals().fetchAll().then(function (principals) {
                        expect(principals).toBeDefined();
                        expect(principals).toEqual(jasmine.any(Array));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("servicePrincipal.spec.2.1 should be able to get service principals (tries to add a servicePrincipal first)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals')).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (created) {
                            that.client.servicePrincipals.getServicePrincipals().fetchAll().then(function (servicePrincipals) {
                                expect(servicePrincipals).toBeDefined();
                                expect(servicePrincipals).toEqual(jasmine.any(Array));
                                expect(servicePrincipals.length).toBeGreaterThan(0);
                                expect(servicePrincipals[0]).toEqual(jasmine.any(Microsoft.AADGraph.ServicePrincipal));
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("servicePrincipal.spec.3 should be able to apply filter to service principals", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals')).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (created) {
                            that.client.servicePrincipals.getServicePrincipals().filter("appId eq '" + created.appId + "'").fetchAll().then(function (servicePrincipals) {
                                expect(servicePrincipals).toBeDefined();
                                expect(servicePrincipals).toEqual(jasmine.any(Array));
                                expect(servicePrincipals.length).toEqual(1);
                                expect(servicePrincipals[0]).toEqual(jasmine.any(Microsoft.AADGraph.ServicePrincipal));
                                expect(servicePrincipals[0].appId).toEqual(created.appId);
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("servicePrincipal.spec.4 should be able to apply top query to service principals", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals1')).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals2')).then(function (app2) {
                            that.tempEntities.push(app2);
                            that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (created) {
                                var secondPrincipal = that.createServicePrincipal(app2);
                                that.client.servicePrincipals.addServicePrincipal(secondPrincipal).then(function (created2) {
                                    that.client.servicePrincipals.getServicePrincipals().top(1).fetch().then(function (servicePrincipals) {
                                        expect(servicePrincipals).toBeDefined();
                                        expect(servicePrincipals.currentPage).toBeDefined();
                                        expect(servicePrincipals.currentPage).toEqual(jasmine.any(Array));
                                        expect(servicePrincipals.currentPage.length).toEqual(1);
                                        expect(servicePrincipals.currentPage[0]).toEqual(jasmine.any(Microsoft.AADGraph.ServicePrincipal));
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("servicePrincipal.spec.5 should be able to get a newly created service principal by objectId", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals')).then(function (app) {
                        that.tempEntities.push(app);
                        var newPrincipal = that.createServicePrincipal(app);
                        that.client.servicePrincipals.addServicePrincipal(newPrincipal).then(function (added) {
                            that.client.servicePrincipals.getServicePrincipal(added.objectId).fetch().then(function (got) {
                                expect(got.objectId).toEqual(added.objectId);
                                expect(got).toEqual(jasmine.any(Microsoft.AADGraph.ServicePrincipal));
                                expect(got.appId).toEqual(added.appId);
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("servicePrincipal.spec.6 should be able to modify existing service principal", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals')).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (added) {
                            var tag = 'TAG';
                            added.tags = [tag];
                            added.update().then(function () {
                                that.client.servicePrincipals.getServicePrincipal(added.objectId).fetch().then(function (got) {
                                    expect(got.objectId).toEqual(added.objectId);
                                    expect(got.appId).toEqual(added.appId);
                                    expect(got.tags.length).toBe(1);
                                    expect(got.tags).toContain(tag);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("servicePrincipal.spec.7 should be able to delete existing service principal", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('AppToTestServicePrincipals')).then(function (app) {
                        that.tempEntities.push(app);
                        testAppId = app.appId;
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (added) {
                            added.delete().then(function () {
                                that.client.servicePrincipals.getServicePrincipal(added.objectId).fetch().then(function (got) {
                                    expect(got).toBeUndefined();
                                    done();
                                }, function (err) {
                                    expect(err.statusText).toBeDefined();
                                    expect(err.statusText).toMatch("Not Found");
                                    done();
                                });
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("OAuth2PermissionGrant", function () {
            it("oauth2permissiongrant.spec.1 should be able to create a new grant", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (app1) {
                        that.tempEntities.push(app1);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app1)).then(function (principal1) {
                            that.client.applications.addApplication(that.createApp('app2')).then(function (app2) {
                                that.tempEntities.push(app2);
                                that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                    that.client.users.addUser(that.createUser()).then(function (user) {
                                        that.tempEntities.push(user);
                                        var newGrant = that.createGrant(principal1.objectId, principal2.objectId, user.objectId);
                                        that.client.oauth2PermissionGrants.addOAuth2PermissionGrant(newGrant).then(function (added) {
                                            expect(added.objectId).toBeDefined();
                                            expect(added.principalId).toMatch(newGrant.principalId);
                                            expect(added.resourceId).toMatch(newGrant.resourceId);
                                            expect(added.clientId).toMatch(newGrant.clientId);
                                            expect(added.path).toMatch(added.objectId);
                                            expect(added).toEqual(jasmine.any(Microsoft.AADGraph.OAuth2PermissionGrant));
                                            done();
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("oauth2permissiongrant.spec.2 should be able to get grants", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.oauth2PermissionGrants.getOAuth2PermissionGrants().fetchAll().then(function (grants) {
                        expect(grants).toBeDefined();
                        expect(grants).toEqual(jasmine.any(Array));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("oauth2permissiongrant.spec.2.1 should be able to get grants (tries to add a grant first)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (app1) {
                        that.tempEntities.push(app1);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app1)).then(function (principal1) {
                            that.client.applications.addApplication(that.createApp('app2')).then(function (app2) {
                                that.tempEntities.push(app2);
                                that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                    that.client.users.addUser(that.createUser()).then(function (user) {
                                        that.tempEntities.push(user);
                                        var newGrant = that.createGrant(principal1.objectId, principal2.objectId, user.objectId);
                                        that.client.oauth2PermissionGrants.addOAuth2PermissionGrant(newGrant).then(function () {
                                            that.client.oauth2PermissionGrants.getOAuth2PermissionGrants().fetchAll().then(function (grants) {
                                                expect(grants).toBeDefined();
                                                expect(grants).toEqual(jasmine.any(Array));
                                                expect(grants.length).toBeGreaterThan(0);
                                                expect(grants[0]).toEqual(jasmine.any(Microsoft.AADGraph.OAuth2PermissionGrant));
                                                done();
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("oauth2permissiongrant.spec.3 should be able to apply filter to grants", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (app1) {
                        that.tempEntities.push(app1);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app1)).then(function (principal1) {
                            that.client.applications.addApplication(that.createApp('app2')).then(function (app2) {
                                that.tempEntities.push(app2);
                                that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                    that.client.users.addUser(that.createUser()).then(function (user) {
                                        that.tempEntities.push(user);
                                        var newGrant = that.createGrant(principal1.objectId, principal2.objectId, user.objectId);
                                        that.client.oauth2PermissionGrants.addOAuth2PermissionGrant(newGrant).then(function (added) {
                                            that.client.oauth2PermissionGrants.getOAuth2PermissionGrants().filter("principalId eq'" + user.objectId + "'").fetchAll().then(function (grants) {
                                                expect(grants).toBeDefined();
                                                expect(grants).toEqual(jasmine.any(Array));
                                                expect(grants.length).toEqual(1);
                                                expect(grants[0]).toEqual(jasmine.any(Microsoft.AADGraph.OAuth2PermissionGrant));
                                                expect(grants[0].objectId).toEqual(added.objectId);
                                                done();
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("oauth2permissiongrant.spec.4 should be able to get a newly created grant by objectId", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (app1) {
                        that.tempEntities.push(app1);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app1)).then(function (principal1) {
                            that.client.applications.addApplication(that.createApp('app2')).then(function (app2) {
                                that.tempEntities.push(app2);
                                that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                    that.client.users.addUser(that.createUser()).then(function (user) {
                                        that.tempEntities.push(user);
                                        var newGrant = that.createGrant(principal1.objectId, principal2.objectId, user.objectId);
                                        that.client.oauth2PermissionGrants.addOAuth2PermissionGrant(newGrant).then(function (added) {
                                            that.client.oauth2PermissionGrants.getOAuth2PermissionGrant(added.objectId).fetch().then(function (got) {
                                                expect(got.objectId).toEqual(added.objectId);
                                                expect(got).toEqual(jasmine.any(Microsoft.AADGraph.OAuth2PermissionGrant));
                                                expect(got.objectId).toEqual(added.objectId);
                                                expect(got.principalId).toEqual(added.principalId);
                                                expect(got.clientId).toEqual(added.clientId);
                                                expect(got.resourceId).toEqual(added.resourceId);
                                                done();
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("oauth2permissiongrant.spec.5 should be able to modify existing grant", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (app1) {
                        that.tempEntities.push(app1);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app1)).then(function (principal1) {
                            that.client.applications.addApplication(that.createApp('app2')).then(function (app2) {
                                that.tempEntities.push(app2);
                                that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                    that.client.users.addUser(that.createUser()).then(function (user) {
                                        that.tempEntities.push(user);
                                        var newGrant = that.createGrant(principal1.objectId, principal2.objectId, user.objectId);
                                        that.client.oauth2PermissionGrants.addOAuth2PermissionGrant(newGrant).then(function (added) {
                                            added.expiryTime = '2014-05-04T00:00:00';
                                            added.update().then(function () {
                                                that.client.oauth2PermissionGrants.getOAuth2PermissionGrant(added.objectId).fetch().then(function (got) {
                                                    expect(got.objectId).toEqual(added.objectId);
                                                    var addedDate = new Date(added.expiryTime);
                                                    expect(got.expiryTime).toEqual(addedDate);
                                                    done();
                                                }, fail.bind(that, done));
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("oauth2permissiongrant.spec.6 should be able to delete existing grant", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (app1) {
                        that.tempEntities.push(app1);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app1)).then(function (principal1) {
                            that.client.applications.addApplication(that.createApp('app2')).then(function (app2) {
                                that.tempEntities.push(app2);
                                that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                    that.client.users.addUser(that.createUser()).then(function (user) {
                                        that.tempEntities.push(user);
                                        var newGrant = that.createGrant(principal1.objectId, principal2.objectId, user.objectId);
                                        that.client.oauth2PermissionGrants.addOAuth2PermissionGrant(newGrant).then(function (added) {
                                            added.delete().then(function () {
                                                that.client.oauth2PermissionGrants.getOAuth2PermissionGrant(added.objectId).fetch().then(function (got) {
                                                    expect(got).toBeUndefined();
                                                    done();
                                                }, function (err) {
                                                    expect(err.statusText).toBeDefined();
                                                    expect(err.statusText).toMatch("Not Found");
                                                    done();
                                                });
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("oauth2permissiongrant.spec.7 should be able to get principal grants", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp('app1')).then(function (app1) {
                        that.tempEntities.push(app1);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app1)).then(function (principal1) {
                            that.client.applications.addApplication(that.createApp('app2')).then(function (app2) {
                                that.tempEntities.push(app2);
                                that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                    that.client.users.addUser(that.createUser()).then(function (user) {
                                        that.tempEntities.push(user);
                                        var newGrant = that.createGrant(principal1.objectId, principal2.objectId, user.objectId);
                                        that.client.oauth2PermissionGrants.addOAuth2PermissionGrant(newGrant).then(function (added) {
                                            that.client.users.getUser(user.objectId).fetch().then(function (updatedUser) {
                                                updatedUser.oauth2PermissionGrants.getOAuth2PermissionGrants().fetchAll().then(function (grants) {
                                                    expect(grants[0].objectId).toBe(added.objectId);
                                                    expect(grants[0].principalId).toBe(added.principalId);
                                                    expect(grants[0].clientId).toBe(added.clientId);
                                                    expect(grants[0].resourceId).toBe(added.resourceId);
                                                    done();
                                                }, fail.bind(that, done));
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("DirectoryRoleTemplate", function () {
            it("dirRoleTemplate.spec.1 should be able to get all directoryRoleTemplates", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.directoryObjects.asDirectoryRoleTemplates().fetchAll().then(function (directoryRoleTemplates) {
                        expect(directoryRoleTemplates).toBeDefined();
                        expect(directoryRoleTemplates).toEqual(jasmine.any(Array));
                        expect(directoryRoleTemplates.length).toBeGreaterThan(0);//exists default directoryRoleTemplates
                        expect(directoryRoleTemplates[0]).toEqual(jasmine.any(Microsoft.AADGraph.DirectoryRoleTemplate));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Devices", function () {
            it("devices.spec.1 should be able to create a new device", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.devices.addDevice(that.createDevice()).then(function (added) {
                        that.tempEntities.push(added);
                        expect(added.objectId).toBeDefined();
                        expect(added.deviceId).toBeDefined();
                        expect(added.path).toMatch(added.objectId);
                        expect(added).toEqual(jasmine.any(Microsoft.AADGraph.Device));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("devices.spec.2 should be able to get devices", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.devices.getDevices().fetchAll().then(function (devices) {
                        expect(devices).toBeDefined();
                        expect(devices).toEqual(jasmine.any(Array));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });

            it("devices.spec.2.1 should be able to get devices (tries to add a device first)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.devices.addDevice(that.createDevice()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.devices.getDevices().fetchAll().then(function (devices) {
                            expect(devices).toBeDefined();
                            expect(devices).toEqual(jasmine.any(Array));
                            expect(devices.length).toBeGreaterThan(0);
                            expect(devices[0]).toEqual(jasmine.any(Microsoft.AADGraph.Device));
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("devices.spec.3 should be able to apply filter to devices", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.devices.addDevice(that.createDevice()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.devices.getDevices().filter("displayName eq '" + created.displayName + "'").fetchAll().then(function (devices) {
                            expect(devices).toBeDefined();
                            expect(devices).toEqual(jasmine.any(Array));
                            expect(devices.length).toEqual(1);
                            expect(devices[0]).toEqual(jasmine.any(Microsoft.AADGraph.Device));
                            expect(devices[0].displayName).toEqual(created.displayName);
                            expect(devices[0].deviceId).toEqual(created.deviceId);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("devices.spec.4 should be able to apply top query to devices", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.devices.addDevice(that.createDevice()).then(function (created) {
                        that.tempEntities.push(created);
                        that.client.devices.addDevice(that.createDevice()).then(function (created2) {
                            that.tempEntities.push(created2);
                            that.client.devices.getDevices().top(1).fetch().then(function (devices) {
                                expect(devices).toBeDefined();
                                expect(devices.currentPage).toBeDefined();
                                expect(devices.currentPage).toEqual(jasmine.any(Array));
                                expect(devices.currentPage.length).toEqual(1);
                                expect(devices.currentPage[0]).toEqual(jasmine.any(Microsoft.AADGraph.Device));
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("devices.spec.5 should be able to get a newly created device by objectId", function (done) {
                var that = this;
                var newUser = that.createDevice();

                that.runSafely(function () {
                    that.client.devices.addDevice(newUser).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.devices.getDevice(added.objectId).fetch().then(function (got) {
                            expect(got.objectId).toEqual(added.objectId);
                            expect(got).toEqual(jasmine.any(Microsoft.AADGraph.Device));
                            expect(got.displayName).toEqual(newUser.displayName);
                            expect(got.deviceId).toEqual(newUser.deviceId);
                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("devices.spec.6 should be able to modify existing device", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.devices.addDevice(that.createDevice()).then(function (added) {
                        that.tempEntities.push(added);
                        added.displayName = guid();
                        added.update().then(function () {
                            that.client.devices.getDevice(added.objectId).fetch().then(function (got) {
                                expect(got.objectId).toEqual(added.objectId);
                                expect(got.displayName).toEqual(added.displayName);
                                expect(got.deviceId).toEqual(added.deviceId);
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("devices.spec.7 should be able to delete existing device", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.devices.addDevice(that.createDevice()).then(function (added) {
                        added.delete().then(function () {
                            that.client.devices.getDevice(added.objectId).fetch().then(function (got) {
                                expect(got).toBeUndefined();
                                done();
                            }, function (err) {
                                expect(err.statusText).toBeDefined();
                                expect(err.statusText).toMatch("Not Found");
                                done();
                            });
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("AppRoleAssignment", function () {
            it("appRoleAssignments.spec.1 should be able to create a new appRoleAssignment", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (principal) {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                user.appRoleAssignments.addAppRoleAssignment(that.createAppRoleAssignment(principal, user)).then(function (added) {
                                    expect(added.objectId).toBeDefined();
                                    expect(added.path).toMatch(added.objectId);
                                    expect(added).toEqual(jasmine.any(Microsoft.AADGraph.AppRoleAssignment));
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("appRoleAssignments.spec.2 should be able to get appRoleAssignments", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (principal) {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                user.appRoleAssignments.addAppRoleAssignment(that.createAppRoleAssignment(principal, user)).then(function (added) {
                                    user.appRoleAssignments.getAppRoleAssignments().fetchAll().then(function (appRoleAssignments) {
                                        expect(appRoleAssignments).toBeDefined();
                                        expect(appRoleAssignments).toEqual(jasmine.any(Array));
                                        expect(appRoleAssignments.length).toBeGreaterThan(0);
                                        expect(appRoleAssignments[0]).toEqual(jasmine.any(Microsoft.AADGraph.AppRoleAssignment));
                                        expect(appRoleAssignments[0].resourceId).toEqual(added.resourceId);
                                        expect(appRoleAssignments[0].principalId).toEqual(added.principalId);
                                        expect(appRoleAssignments[0].id).toEqual(added.id);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("appRoleAssignments.spec.3 should be able to apply filter to appRoleAssignments", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (principal) {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                user.appRoleAssignments.addAppRoleAssignment(that.createAppRoleAssignment(principal, user)).then(function (added) {
                                    user.appRoleAssignments.getAppRoleAssignments().filter("objectId eq '" + added.objectId + "'").fetchAll().then(function (appRoleAssignments) {
                                        expect(appRoleAssignments).toBeDefined();
                                        expect(appRoleAssignments).toEqual(jasmine.any(Array));
                                        expect(appRoleAssignments.length).toEqual(1);
                                        expect(appRoleAssignments[0]).toEqual(jasmine.any(Microsoft.AADGraph.AppRoleAssignment));
                                        expect(appRoleAssignments[0].objectId).toEqual(added.objectId);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("appRoleAssignments.spec.4 should be able to apply top query to appRoleAssignments", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (principal) {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                user.appRoleAssignments.addAppRoleAssignment(that.createAppRoleAssignment(principal, user)).then(function (added) {
                                    that.client.applications.addApplication(that.createApp('AppToTest2')).then(function (app2) {
                                        that.tempEntities.push(app2);
                                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app2)).then(function (principal2) {
                                            user.appRoleAssignments.addAppRoleAssignment(that.createAppRoleAssignment(principal2, user)).then(function (added2) {
                                                user.appRoleAssignments.getAppRoleAssignments().top(1).fetch().then(function (appRoleAssignments) {
                                                    expect(appRoleAssignments).toBeDefined();
                                                    expect(appRoleAssignments.currentPage).toBeDefined();
                                                    expect(appRoleAssignments.currentPage).toEqual(jasmine.any(Array));
                                                    expect(appRoleAssignments.currentPage.length).toEqual(1);
                                                    expect(appRoleAssignments.currentPage[0]).toEqual(jasmine.any(Microsoft.AADGraph.AppRoleAssignment));
                                                    done();
                                                }, fail.bind(that, done));
                                            }, fail.bind(that, done));
                                        }, fail.bind(that, done));
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("appRoleAssignments.spec.5 should be able to get a newly created appRoleAssignment by objectId", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        that.client.servicePrincipals.addServicePrincipal(that.createServicePrincipal(app)).then(function (principal) {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                that.tempEntities.push(user);
                                user.appRoleAssignments.addAppRoleAssignment(that.createAppRoleAssignment(principal, user)).then(function (added) {
                                    user.appRoleAssignments.getAppRoleAssignment(added.objectId).fetch().then(function (got) {
                                        expect(got.objectId).toEqual(added.objectId);
                                        expect(got).toEqual(jasmine.any(Microsoft.AADGraph.AppRoleAssignment));
                                        expect(got.objectId).toEqual(added.objectId);
                                        expect(got.resourceId).toEqual(added.resourceId);
                                        expect(got.principalId).toEqual(added.principalId);
                                        expect(got.id).toEqual(added.id);
                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Paging", function () {
            it("paging.spec.1 should be able to get users using nextPage", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.addANumberOfUsers(105).then(function (added) {
                        that.tempEntities = added;
                        that.client.users.getUsers().fetch().then(function (obj) {
                            expect(obj.currentPage).toBeDefined();
                            expect(obj.currentPage).toEqual(jasmine.any(Array));
                            expect(obj.currentPage.length).toBe(100);
                            obj.getNextPage().then(function (obj2) {
                                expect(obj2.currentPage).toBeDefined();
                                expect(obj2.currentPage).toEqual(jasmine.any(Array));
                                // As some other objects could exist there
                                expect(obj2.currentPage.length >= 5).toBeTruthy();
                                expect(obj2.currentPage.length <= 100).toBeTruthy();
                                for (var i = 0; i < obj2.currentPage.length; i++) {
                                    expect(obj.currentPage).not.toContainObjWithId(obj2.currentPage[i].objectId);
                                }
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("paging.spec.2 should be able to get users using nextPage (top)", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.addANumberOfUsers(105).then(function (added) {
                        that.tempEntities = added;
                        that.client.users.getUsers().top(101).fetch().then(function (obj) {
                            expect(obj.currentPage).toBeDefined();
                            expect(obj.currentPage).toEqual(jasmine.any(Array));
                            expect(obj.currentPage.length).toBe(101);
                            obj.getNextPage().then(function (obj2) {
                                expect(obj2.currentPage).toBeDefined();
                                expect(obj2.currentPage).toEqual(jasmine.any(Array));
                                expect(obj2.currentPage.length >= 4).toBeTruthy();
                                expect(obj2.currentPage.length <= 100).toBeTruthy();
                                for (var i = 0; i < obj2.currentPage.length; i++) {
                                    expect(obj.currentPage).not.toContainObjWithId(obj2.currentPage[i].objectId);
                                }
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("paging.spec.3 should be able to get users using nextPage (filter + fetch)", function (done) {
                var that = this;
                var prefixForUserName = 'testPagingspec3';

                that.runSafely(function () {
                    that.addANumberOfUsers(105, prefixForUserName).then(function (added) {
                        that.tempEntities = added;
                        that.client.users.getUsers().filter("startswith(displayName,'" + prefixForUserName + "')").fetch().then(function (obj) {
                            expect(obj.currentPage).toBeDefined();
                            expect(obj.currentPage).toEqual(jasmine.any(Array));
                            expect(obj.currentPage.length).toBe(100);
                            expect(obj.currentPage[0].displayName.indexOf(prefixForUserName)).not.toEqual(-1);
                            obj.getNextPage().then(function (obj2) {
                                expect(obj2.currentPage).toBeDefined();
                                expect(obj2.currentPage).toEqual(jasmine.any(Array));
                                expect(obj2.currentPage.length).toEqual(5);
                                expect(obj2.currentPage[0].displayName).toContain(prefixForUserName);
                                for (var i = 0; i < obj2.currentPage.length; i++) {
                                    expect(obj.currentPage).not.toContainObjWithId(obj2.currentPage[i].objectId);
                                }
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Orderby and combined queries", function () {
            it("queries.spec.1 should be able to use orderBy", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser('TestZ')).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.users.addUser(that.createUser('TestA')).then(function (added2) {
                            that.tempEntities.push(added2);
                            that.client.users.getUsers().orderBy('displayName').fetchAll().then(function (users) {
                                expect(users).toBeDefined();
                                expect(users).toEqual(jasmine.any(Array));
                                expect(users[0].displayName).toBeLessThan(users[1].displayName);
                                done();
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
            it("queries.spec.2 should be able to use orderBy + top", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser('TestZ')).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.users.addUser(that.createUser('TestA')).then(function (added2) {
                            that.tempEntities.push(added2);
                            that.client.users.addUser(that.createUser('TestB')).then(function (added3) {
                                that.tempEntities.push(added3);
                                that.client.users.getUsers().orderBy('displayName').top(2).fetch().then(function (users) {
                                    expect(users.currentPage).toBeDefined();
                                    expect(users.currentPage).toEqual(jasmine.any(Array));
                                    expect(users.currentPage.length).toEqual(2);
                                    expect(users.currentPage[0].displayName).toBeLessThan(users.currentPage[1].displayName);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
            it("queries.spec.3 should be able to use filter + top", function (done) {
                var that = this;
                var prefixName = 'TestPrefix';

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser(prefixName + 'A')).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.users.addUser(that.createUser(prefixName + 'B')).then(function (added2) {
                            that.tempEntities.push(added2);
                            that.client.users.addUser(that.createUser(prefixName + 'C')).then(function (added3) {
                                that.tempEntities.push(added3);
                                that.client.users.getUsers().filter("startswith(displayName,'" + prefixName + "')").top(2).fetch().then(function (users) {
                                    expect(users.currentPage).toBeDefined();
                                    expect(users.currentPage).toEqual(jasmine.any(Array));
                                    expect(users.currentPage.length).toEqual(2);
                                    expect(users.currentPage[0].displayName).toContain(prefixName);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
            it("queries.spec.3.1 should be able to use filter + top when top argument is too large", function (done) {
                var that = this;
                var prefixName = 'TestPrefix';

                that.runSafely(function () {
                    that.client.users.addUser(that.createUser(prefixName + 'A')).then(function (added) {
                        that.tempEntities.push(added);
                        that.client.users.addUser(that.createUser(prefixName + 'B')).then(function (added2) {
                            that.tempEntities.push(added2);
                            that.client.users.addUser(that.createUser(prefixName + 'C')).then(function (added3) {
                                that.tempEntities.push(added3);
                                that.client.users.getUsers().filter("startswith(displayName,'" + prefixName + "')").top(4).fetch().then(function (users) {
                                    expect(users.currentPage).toBeDefined();
                                    expect(users.currentPage).toEqual(jasmine.any(Array));
                                    expect(users.currentPage.length).toEqual(3);
                                    expect(users.currentPage[0].displayName).toContain(prefixName);
                                    expect(users.currentPage[users.currentPage.length - 1].displayName).toContain(prefixName);
                                    done();
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Contacts", function () {
            it("contacts.spec.1 should be able to get contacts", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.contacts.getContacts().fetchAll().then(function (contacts) {
                        expect(contacts).toBeDefined();
                        expect(contacts).toEqual(jasmine.any(Array));
                        // This needs a real synced Active Directory with contacts
                        //expect(contacts.length).toBeGreaterThan(0);
                        //expect(contacts[0]).toEqual(jasmine.any(Microsoft.AADGraph.Contact));
                        done();
                    }, fail.bind(that, done));
                }, done);
            });
        });

        describe("Directory Schema Extensions", function () {
            it("extensions.spec.1 should be able to add an extension", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        var extensionRaw = that.createExtensionProperty('skypeId');

                        app.extensionProperties.addExtensionProperty(extensionRaw).then(function (extension) {
                            // No need to remove the extension as we will remove the app after the test
                            expect(extension).toBeDefined();
                            expect(extension).toEqual(jasmine.any(Microsoft.AADGraph.ExtensionProperty));
                            expect(extension.objectId).toBeDefined();
                            // The name should become a fully-qualified extension property name
                            expect(extension.name).not.toEqual(extensionRaw.name);

                            done();
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("extensions.spec.2 should be able to get registered extensions", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        var extensionRaw = that.createExtensionProperty('skypeId');

                        app.extensionProperties.addExtensionProperty(extensionRaw).then(function (extension) {
                            app.extensionProperties.getExtensionProperties().fetch().then(function (extensions) {
                                expect(extensions).toBeDefined();
                                expect(extensions.currentPage).toEqual(jasmine.any(Array));
                                expect(extensions.currentPage).toContainObjWithId(extension.objectId);

                                done();
                            });
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            xit("extensions.spec.3 should be able to write a registered extension value", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        var extensionRaw = that.createExtensionProperty('skypeId');
                        var extensionValue = 'someskypename';

                        app.extensionProperties.addExtensionProperty(extensionRaw).then(function (extension) {
                            that.client.users.addUser(that.createUser()).then(function (user) {
                                user[extension.name] = extensionValue;

                                // Fails with: {"odata.error":{"code":"Request_BadRequest","message":{"lang":"en","value":"The extension property(s) extension_97b0c9e518424a31a5ec916e286f8a49_skypeId is\/are not available."},"values":null}}
                                user.update().then(function () {
                                    that.client.users.getUser(user.objectId).fetch().then(function (updatedUser) {
                                        expect(updatedUser).toBeDefined();
                                        expect(updatedUser[extension.name]).toEqual(extensionValue);

                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });

            it("extensions.spec.4 should be able to remove a registered extension", function (done) {
                var that = this;

                that.runSafely(function () {
                    that.client.applications.addApplication(that.createApp()).then(function (app) {
                        that.tempEntities.push(app);
                        var extensionRaw = that.createExtensionProperty('skypeId');

                        app.extensionProperties.addExtensionProperty(extensionRaw).then(function (extension) {
                            app.extensionProperties.getExtensionProperties().fetch().then(function (extensions) {
                                expect(extensions).toBeDefined();
                                expect(extensions.currentPage).toEqual(jasmine.any(Array));
                                expect(extensions.currentPage).toContainObjWithId(extension.objectId);

                                extension.delete().then(function () {
                                    app.extensionProperties.getExtensionProperties().fetch().then(function (extensionsAfterRemove) {
                                        expect(extensionsAfterRemove).toBeDefined();
                                        expect(extensionsAfterRemove.currentPage).toEqual(jasmine.any(Array));
                                        expect(extensionsAfterRemove.currentPage).not.toContainObjWithId(extension.objectId);

                                        done();
                                    }, fail.bind(that, done));
                                }, fail.bind(that, done));
                            }, fail.bind(that, done));
                        }, fail.bind(that, done));
                    }, fail.bind(that, done));
                }, done);
            });
        });
    });
};

exports.defineManualTests = function (contentEl, createActionButton) {
    var authContext;

    createActionButton('Log in', function () {
        authContext = new AuthenticationContext(AUTH_URL);
        authContext.acquireTokenAsync(RESOURCE_URL, APP_ID, REDIRECT_URL).then(function (authRes) {
            // Save acquired userId for further usage
            TEST_USER_ID = authRes.userInfo && authRes.userInfo.userId;

            console.log("Token is: " + authRes.accessToken);
            console.log("TEST_USER_ID is: " + TEST_USER_ID);
        }, function (err) {
            console.error(err);
        });
    });

    createActionButton('Log out', function () {
        authContext = authContext || new AuthenticationContext(AUTH_URL);
        return authContext.tokenCache.clear().then(function () {
            console.log("Logged out");
        }, function (err) {
            console.error(err);
        });
    });
};
