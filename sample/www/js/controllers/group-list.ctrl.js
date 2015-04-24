
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
    'use strict';

    angular
      .module('starter.controllers')
      .controller('GroupListCtrl', ['$state', '$rootScope', '$ionicModal', '$scope', 'AadClient', GroupListCtrl]);
    
    function GroupListCtrl($state, $rootScope, $ionicModal, $scope, AadClient) {
        var vm = this;
        vm.open = open;
        vm.create = create;
        vm.cancelModal = cancelModal;

        initModal();
        activate();

        $rootScope.$on('groups:listChanged', activate);

        return vm;

        ///////////

        function initModal() {
            $ionicModal.fromTemplateUrl('views/group-create.html', {
                scope: $scope
            }).then(function (modal) {
                vm.modal = modal;
            });
        }

        function activate() {
            $scope.showLoading();
            AadClient.getGroups().then(function (groups) {
                vm.groups = groups;
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function open(groupId) {
            $state.go('app.group-detail', { 'groupId': groupId });
        }

        function create(createForm, newGroup) {
            if (createForm.$valid) {
                $scope.showLoading();
                AadClient.addGroup(newGroup.name, newGroup.description).then(function() {
                    vm.modal.hide();
                    initModal();
                    $scope.hideLoading();
                    $scope.$emit('groups:listChanged');
                }, $scope.errorHandler);
            }
        }

        function cancelModal() {
            vm.modal.hide();
            initModal();
        }
    }
})();