
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
    'use strict';

    angular
      .module('starter.controllers')
      .controller('AppDetailCtrl', ['$state', '$scope', '$stateParams', '$ionicPopup', '$ionicHistory', '$ionicModal', 'InputHandler','CloneManager', 'AadClient', AppDetailCtrl]);

    function AppDetailCtrl($state, $scope, $stateParams, $ionicPopup, $ionicHistory, $ionicModal, InputHandler, CloneManager, AadClient) {
        var vm = this;
        vm.edit = edit;
        vm.cancelModal = cancelModal;
        vm.updatedFields = { identifierUris: '', displayName: '' };

        initEditModal();
        activate();

        vm.showAppDeleteConfirmation = showAppDeleteConfirmation;

        return vm;

        //////////////

        function activate() {
            $scope.showLoading();
            AadClient.getApp($stateParams.objectId).then(function (app) {
                vm.app = app;
                initializeUpdatedFields(vm.app);
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function initEditModal() {
            $ionicModal.fromTemplateUrl('views/application-edit.html', {
                scope: $scope
            }).then(function (editModal) {
                vm.editModal = editModal;
            });
        }

        function showAppDeleteConfirmation() {
            var confirmPopup = $ionicPopup.confirm({
                title: 'Confirm deletion',
                template: 'Are you sure you want to delete this application?',
                okText: 'Delete',
                okType: 'button-assertive'
            });
            confirmPopup.then(function (res) {
                if (res) {
                    remove();
                }
            });
        }

        function remove() {
            $scope.showLoading();
            AadClient.deleteApp($stateParams.objectId).then(function () {
                $ionicHistory.nextViewOptions({
                    disableBack: true
                });
                $scope.hideLoading();
                $scope.$emit('applications:listChanged');
                $scope.$emit('deletedApps:listChanged');
                $state.go('app.application-list');
            }, $scope.errorHandler);
        }

        function edit(editForm) {
            if (editForm.$valid) {
                var cloned = CloneManager.clone(vm.app);
                $scope.showLoading();
                var identifierUris = InputHandler.uriStringToArray(vm.updatedFields.identifierUris);

                AadClient.editApp(cloned, identifierUris, vm.updatedFields.displayName).then(function () {
                    vm.editModal.hide();
                    $scope.hideLoading();
                    initEditModal();
                    $scope.$emit('applications:listChanged');

                    activate();
                }, $scope.errorHandler);
            }
        }

        function cancelModal() {
            vm.editModal.hide();
            initEditModal();
            initializeUpdatedFields(vm.app);
        }

        function initializeUpdatedFields(app) {
            vm.updatedFields.identifierUris = app.identifierUris.join(', ');
            vm.updatedFields.displayName = app.displayName;
        }
    }
})();