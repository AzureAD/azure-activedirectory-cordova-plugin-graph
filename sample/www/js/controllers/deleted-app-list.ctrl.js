
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
    'use strict';

    angular
        .module('starter.controllers')
        .controller('DeletedAppListCtrl', ['$state', '$rootScope', '$ionicModal', '$scope', '$ionicLoading', 'InputHandler', 'AadClient', DeletedAppListCtrl]);

    function DeletedAppListCtrl($state, $rootScope, $ionicModal, $scope, $ionicLoading, InputHandler, AadClient) {
        var vm = this;
        vm.open = open;

        activate();

        $rootScope.$on('deletedApps:listChanged', activate);

        return vm;

        ///////////

        function activate() {
            $scope.showLoading();
            AadClient.getDeletedApps().then(function (deletedApps) {
                vm.deletedApps = deletedApps;
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function open(objectId) {
            $state.go('app.deleted-app-detail', { 'objectId': objectId });
        }
    }

})();