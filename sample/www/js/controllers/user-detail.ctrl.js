
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
    'use strict';

    angular
      .module('starter.controllers')
      .controller('UserDetailCtrl', ['$state', '$scope', '$stateParams', '$ionicPopup', '$ionicHistory', '$ionicModal', 'tenantName','CloneManager', 'AadClient', UserDetailCtrl]);

    function UserDetailCtrl($state, $scope, $stateParams, $ionicPopup, $ionicHistory, $ionicModal, tenantName, CloneManager, AadClient) {
        var vm = this;
        vm.tenantName = tenantName;
        vm.edit = edit;
        vm.cancelEditModal = cancelEditModal;
        vm.cancelResetPasswordModal = cancelResetPasswordModal;

        vm.updatedFields = { userName: '', displayName: '', firstName: '', lastName: '', password: '', passwordCopy: '' };
        vm.onReadOnlyInputChanged = onReadOnlyInputChanged;

        initEditModal();
        initResetPasswordModal();
        activate();

        vm.showDeleteConfirmation = showDeleteConfirmation;
        vm.showResetPasswordConfirmation = showResetPasswordConfirmation;

        return vm;

        //////////////

        function activate() {
            $scope.showLoading();
            AadClient.getUser($stateParams.userId).then(function (user) {
                vm.user = user;
                initializeUpdatedFields(vm.user);
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function initEditModal() {
            $ionicModal.fromTemplateUrl('views/user-edit.html', {
                scope: $scope
            }).then(function (editModal) {
                vm.editModal = editModal;
            });
        }

        function initResetPasswordModal() {
            $ionicModal.fromTemplateUrl('views/user-reset-password.html', {
                scope: $scope
            }).then(function (resetPasswordModal) {
                vm.resetPasswordModal = resetPasswordModal;
            });
        }

        function showDeleteConfirmation() {
            var confirmPopup = $ionicPopup.confirm({
                title: 'Confirm deletion',
                template: 'Are you sure you want to delete this user?',
                okText: 'Delete',
                okType: 'button-assertive'
            });
            confirmPopup.then(function (res) {
                if (res) {
                    remove();
                }
            });
        }

        function showResetPasswordConfirmation() {
            var confirmPopup = $ionicPopup.confirm({
                title: 'Confirm reseting password',
                template: 'Are you sure you want to reset the password?',
                okText: 'Reset password',
                okType: 'button-assertive'
            });
            confirmPopup.then(function (res) {
                if (res) {
                    resetPassword();
                }
            });
        }

        function remove() {
            $scope.showLoading();
            AadClient.deleteUser($stateParams.userId).then(function () {
                $ionicHistory.nextViewOptions({
                    disableBack: true
                });
                $scope.hideLoading();
                $scope.$emit('users:listChanged');
                $state.go('app.user-list');
            }, $scope.errorHandler);
        }

        function edit(editForm) {
            if (editForm.$valid) {                
                $scope.showLoading();
                var cloned = CloneManager.clone(vm.user);
                AadClient.editUser(cloned, vm.updatedFields.userName, vm.updatedFields.displayName, vm.updatedFields.firstName, vm.updatedFields.lastName).then(function () {
                    vm.editModal.hide();
                    $scope.hideLoading();
                    initEditModal();
                    $scope.$emit('users:listChanged');
                    activate();
                }, $scope.errorHandler);
            }
        }

        function resetPassword() {
            $scope.showLoading();
            var cloned = CloneManager.clone(vm.user);
            AadClient.resetPassword(cloned).then(function (temporaryPassword) {
                vm.updatedFields.password = temporaryPassword;
                vm.updatedFields.passwordCopy = temporaryPassword;
                $scope.hideLoading();
                vm.resetPasswordModal.show();
            }, $scope.errorHandler);
        }

        function cancelEditModal() {
            vm.editModal.hide();
            initEditModal();
            initializeUpdatedFields(vm.user);
        }

        function cancelResetPasswordModal() {
            vm.resetPasswordModal.hide();
            initResetPasswordModal();
        }

        function initializeUpdatedFields(user) {
            vm.updatedFields.userName = user.userPrincipalName.substr(0, user.userPrincipalName.indexOf('@'));
            vm.updatedFields.displayName = user.displayName;
            vm.updatedFields.firstName = user.givenName;
            vm.updatedFields.lastName = user.surname;
        }

        function onReadOnlyInputChanged() {
            //Windows Phone 8.1 does not allow copying text in readonly input.
            //The app uses ng-change directive and passwordCopy variable instead.
            vm.updatedFields.password = vm.updatedFields.passwordCopy;//cancel input
        }
    }
})();