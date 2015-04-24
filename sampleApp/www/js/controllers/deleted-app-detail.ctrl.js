
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
    'use strict';

    angular
      .module('starter.controllers')
      .controller('DeletedAppDetailCtrl', ['$state', '$scope', '$stateParams', '$ionicPopup', '$ionicHistory', '$ionicModal', 'InputHandler', 'AadClient', DeletedAppDetailCtrl]);

    function DeletedAppDetailCtrl($state, $scope, $stateParams, $ionicPopup, $ionicHistory, $ionicModal, InputHandler, AadClient) {
        var vm = this;
        vm.restore = restore;
        vm.cancelModal = cancelModal;
        vm.updatedUris = '';

        initRestoreModal();
        activate();

        return vm;

        //////////////

        function activate() {
            $scope.showLoading();
            AadClient.getDeletedApp($stateParams.objectId).then(function (app) {
                vm.app = app;
                initializeUpdatedUris(vm.app);
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function initRestoreModal() {
            $ionicModal.fromTemplateUrl('views/deleted-app-restore.html', {
                scope: $scope
            }).then(function (restoreModal) {
                vm.restoreModal = restoreModal;
            });
        }

        function restore(restoreForm) {
            if (restoreForm.$valid) {
                $scope.showLoading();
                var identifierUris = InputHandler.uriStringToArray(vm.updatedUris);

                AadClient.restoreApp(vm.app, identifierUris).then(function () {
                    vm.restoreModal.hide();
                    $scope.$emit('applications:listChanged');
                    $scope.$emit('deletedApps:listChanged');
                    $scope.hideLoading();
                    $state.go('app.deleted-app-list');
                }, $scope.errorHandler);
            }
        }

        function cancelModal() {
            vm.restoreModal.hide();
            initRestoreModal();
            initializeUpdatedUris(vm.app);
        }

        function initializeUpdatedUris(app) {
            vm.updatedUris = app.identifierUris.join(', ');
        }
    }
})();