(function () {
    'use strict';

    angular
      .module('starter.controllers')
      .controller('UserListCtrl', ['$state', '$rootScope', '$ionicModal', '$scope', '$ionicLoading', 'tenantName', 'CloneManager', 'AadClient', UserListCtrl]);

    function UserListCtrl($state, $rootScope, $ionicModal, $scope, $ionicLoading, tenantName, CloneManager, AadClient) {
        var vm = this;
        vm.tenantName = tenantName;
        vm.open = open;
        vm.create = create;
        vm.cancelAddModal = cancelAddModal;
        vm.cancelTempPasswordModal = cancelTempPasswordModal;

        //for adding new user
        initializeUpdatedFields();
        vm.password = '';
        vm.passwordCopy = '';
        vm.userPrincipalName = '';
        vm.onReadOnlyInputChanged = onReadOnlyInputChanged;

        initAddModal();
        initTempPasswordModal();

        activate();

        $rootScope.$on('users:listChanged', activate);

        return vm;

        ///////////

        function initAddModal() {
            $ionicModal.fromTemplateUrl('views/user-create.html', {
                scope: $scope
            }).then(function (addModal) {
                vm.addModal = addModal;
            });
        }

        function initTempPasswordModal() {
            $ionicModal.fromTemplateUrl('views/user-temporary-password.html', {
                scope: $scope
            }).then(function (tempPasswordModal) {
                vm.tempPasswordModal = tempPasswordModal;
            });
        }

        function activate() {
            $scope.showLoading();
            AadClient.getUsers().then(function (users) {
                vm.users = users;
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function open(userId) {
            $state.go('app.user-detail', { 'userId': userId });
        }

        function create(createForm) {
            if (createForm.$valid) {
                $scope.showLoading();
                var cloned = CloneManager.clone(vm.updatedFields);
                AadClient.addUser(cloned.userName, cloned.displayName, cloned.firstName, cloned.lastName).then(function (addedUser) {
                    vm.addModal.hide();
                    initAddModal();
                    $scope.$emit('users:listChanged');
                    vm.password = addedUser.passwordProfile.password;
                    vm.passwordCopy = addedUser.passwordProfile.password;
                    vm.userPrincipalName = addedUser.userPrincipalName;
                    initializeUpdatedFields();
                    $scope.hideLoading();
                    vm.tempPasswordModal.show();
                }, $scope.errorHandler);
            }
        }

        function cancelAddModal() {
            vm.addModal.hide();
            initializeUpdatedFields();
            initAddModal();
        }

        function cancelTempPasswordModal() {
            vm.tempPasswordModal.hide();
            initTempPasswordModal();
        }

        function initializeUpdatedFields() {
            vm.updatedFields = { userName: '', displayName: '', firstName: '', lastName: ''};
        }

        function onReadOnlyInputChanged() {
            vm.password = vm.passwordCopy;
        }
    }
})();