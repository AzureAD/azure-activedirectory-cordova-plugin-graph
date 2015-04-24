
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
  'use strict';

  angular
    .module('starter.controllers')
      .controller('AppCtrl', ['$scope', '$ionicModal', '$timeout', '$location', '$ionicLoading', '$ionicPopup', 'PromiseFactory', 'AadClient', AppCtrl]);

  function AppCtrl($scope, $ionicModal, $timeout, $location, $ionicLoading, $ionicPopup, PromiseFactory, AadClient) {
        var LOADING_TIMEOUT = 15 * 1000;

        $scope.login = function () {
            AadClient.isLoginRequired().then(function (isRequired) {
                if (isRequired) {
                    $scope.showLoading();
                    AadClient.authenticate().then(function (token) {
                        $scope.$emit('applications:listChanged');
                        $scope.$emit('deletedApps:listChanged');
                        $scope.$emit('users:listChanged');
                        $scope.$emit('groups:listChanged');
                        $scope.hideLoading();
                    }, $scope.errorHandler);
                }
            }, $scope.errorHandler);
        };

        $scope.logOut = function () {
            AadClient.isLoginRequired().then(function (isRequired) {
                $scope.showLoading();
                AadClient.logOut().then(function() {
                    $scope.hideLoading();
                    if (isRequired === false) {
                        $ionicPopup.alert({
                            title: 'Message',
                            template: 'Logged out successfully!',
                        });
                    }
                }, $scope.errorHandler);
            }, $scope.errorHandler);            
        };

        // Open links in side menu
        $scope.sidemenu = function (link) {
          $location.path(link);
        };

        $scope.showLoading = function () {
            $ionicLoading.show({
                template: '<i class="icon ion-loading-c"></i> Loading...',
                duration: LOADING_TIMEOUT
            });
        };

        $scope.hideLoading = function () {
            $ionicLoading.hide();
        };

        $scope.errorHandler = function (err) {
            var promise = PromiseFactory.createPromise();
            var errMessage;

            $scope.hideLoading();

            if (err != null && err.responseText != null && err.responseText != '') {
                errMessage = err.responseText;
            } else if (err != null & err.toString() != '' && err.toString().indexOf('[object') === -1) {
                errMessage = err;
            } else {
                errMessage = 'Unknown error. Please check your network connection.';
            }

            var alertPopup = $ionicPopup.alert({
                title: 'Error occured',
                template: '<i class="icon ion-alert-circled"</i> ' + errMessage
            });

            alertPopup.then(function () {
                promise.resolve();
            });

            return promise;
        };
    }
})();