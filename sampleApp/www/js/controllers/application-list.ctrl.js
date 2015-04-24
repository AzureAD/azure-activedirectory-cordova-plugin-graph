
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function() {
    'use strict';

    angular
        .module('starter.controllers')
        .controller('AppListCtrl', ['$state', '$rootScope', '$ionicModal', '$scope', '$ionicLoading', 'InputHandler','AadClient', AppListCtrl]);

    function AppListCtrl($state, $rootScope, $ionicModal, $scope, $ionicLoading, InputHandler, AadClient) {
        var vm = this;
        vm.open = open;
        vm.create = create;
        vm.cancelModal = cancelModal;
        vm.showDeletedApps = showDeletedApps;

        initModal();
        activate();

        $rootScope.$on('applications:listChanged', activate);

        return vm;

        ///////////

        function initModal() {
            $ionicModal.fromTemplateUrl('views/application-create.html', {
                scope: $scope
            }).then(function (modal) {
                vm.modal = modal;
            });
        }

        function activate() {
            $scope.showLoading();
            AadClient.getApps().then(function (apps) {
                vm.apps = apps;
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function open(objectId) {
            $state.go('app.application-detail', { 'objectId': objectId });
        }

        function create(createForm, appFields) {
            if (createForm.$valid) {
                $scope.showLoading();
                var identifiersUris = InputHandler.uriStringToArray(appFields.identifierUris);

                AadClient.addApp(identifiersUris, appFields.name).then(function() {
                    vm.modal.hide();
                    initModal();
                    $scope.hideLoading();
                    $scope.$emit('applications:listChanged');
                }, $scope.errorHandler);
            }
        }

        function cancelModal() {
            vm.modal.hide();
            initModal();
        }

        function showDeletedApps() {
            $state.go('app.deleted-app-list');
        }
    }

})();