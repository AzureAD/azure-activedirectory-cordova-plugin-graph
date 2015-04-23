(function () {
    'use strict';

    angular
      .module('starter.controllers')
      .controller('GroupDetailCtrl', ['$state', '$ionicPopup', '$ionicHistory', '$ionicModal', '$stateParams', '$scope', 'CloneManager', 'PromiseFactory', 'AadClient', GroupDetailCtrl]);

    function GroupDetailCtrl($state, $ionicPopup, $ionicHistory, $ionicModal, $stateParams, $scope, CloneManager, PromiseFactory, AadClient) {
        var vm = this;
        vm.showGroupDeleteConfirmation = showGroupDeleteConfirmation;
        vm.remove = remove;
        vm.edit = edit;
        vm.cancelEditModal = cancelEditModal;
        vm.open = open;
        vm.deleteMember = deleteMember;
        vm.possibleMembers = [];
        vm.selectMembers = selectMembers;
        vm.updatedFields = { description: '', displayName: '' };
        vm.showAddMember = showAddMember;

        initAddMemberModal();
        initEditModal();
        activate();        

        return vm;

        //////////////

        function initAddMemberModal() {
            var promise = PromiseFactory.createPromise();

            $ionicModal.fromTemplateUrl('views/add-member.html', {
                scope: $scope
            }).then(function (addMemberModal) {
                vm.addMemberModal = addMemberModal;
                promise.resolve();
            });

            return promise;
        }

        function initEditModal() {
            $ionicModal.fromTemplateUrl('views/group-edit.html', {
                scope: $scope
            }).then(function (editModal) {
                vm.editModal = editModal;
            });
        }

        function activate() {
            $scope.showLoading();
            AadClient.getGroup($stateParams.groupId).then(function (group) {
                vm.group = group;
                initializeUpdatedFields(group);
                activateMembers();                
            }, $scope.errorHandler);
        }

        // TODO: change this to update event propagation
        function activateMembers() {
            $scope.showLoading();
            AadClient.getGroupMembers($stateParams.groupId).then(function (members) {
                vm.members = members;
                $scope.hideLoading();
            }, $scope.errorHandler);
        }

        function activatePossibleMembers() {
            var promise = PromiseFactory.createPromise();

            $scope.showLoading();
            AadClient.getPossibleGroupMembers($stateParams.groupId).then(function (possibleMembers) {
                vm.possibleMembers = possibleMembers;
                $scope.hideLoading();
                initAddMemberModal().then(function() {
                    promise.resolve();
                }, promise.reject);
            }, promise.reject);

            return promise;
        }

        function showGroupDeleteConfirmation() {
            var confirmPopup = $ionicPopup.confirm({
                title: 'Confirm deletion',
                template: 'Are you sure you want to delete this group?',
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
            AadClient.deleteGroup($stateParams.groupId).then(function () {
                $ionicHistory.nextViewOptions({
                    disableBack: true
                });
                $scope.hideLoading();
                $scope.$emit('groups:listChanged');
                $state.go('app.group-list');
            }, $scope.errorHandler);
        }

        function edit(editForm) {
            if (editForm.$valid) {
                var cloned = CloneManager.clone(vm.group);
                $scope.showLoading();
                AadClient.editGroup(cloned, vm.updatedFields.displayName, vm.updatedFields.description).then(function() {
                    vm.editModal.hide();
                    initEditModal();
                    $scope.hideLoading();
                    $scope.$emit('groups:listChanged');

                    $scope.showLoading();
                    AadClient.getGroup($stateParams.groupId).then(function(group) {
                        vm.group = group;
                        initializeUpdatedFields(group);
                        $scope.hideLoading();
                    });
                });
            }
        }

        function cancelEditModal() {
            vm.editModal.hide();
            initEditModal();
            initializeUpdatedFields(vm.group);
        }

        function deleteMember(memberId) {
            $scope.showLoading();
            AadClient.deleteGroupMember($stateParams.groupId, memberId).then(function() {
                $scope.hideLoading();
                activateMembers();
            }, $scope.errorHandler);
        }

        function open(groupId) {
            $state.go('app.group-detail', { 'groupId': groupId });
        }

        function showAddMember() {
            activatePossibleMembers().then(function() {
                vm.addMemberModal.show();
            }, $scope.errorHandler);
        }

        function selectMembers() {
            vm.addMemberModal.hide();

            var selectedMembers = vm.possibleMembers.filter(function(item) {
                return item.checked === true;
            });

            if (selectedMembers.length > 0) {
            $scope.showLoading();
                AadClient.addGroupMembers($stateParams.groupId, selectedMembers).then(function() {
                $scope.hideLoading();

                activateMembers();
            }, $scope.errorHandler);
        }
        }

        function initializeUpdatedFields(group) {
            vm.updatedFields.displayName = group.displayName;
            vm.updatedFields.description = group.description;
        }
    }
})();