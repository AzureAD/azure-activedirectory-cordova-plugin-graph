
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function () {
    'use strict';

  // TODO: Investigate why this code started to cause crashes on Windows 8.0/8.1
  //   // This prevents event click to be called twice on Windows 8.1
  //   window.addEventListener('click', function (event) {
  //       if (typeof event.target !== 'undefined' && typeof event.target.attributes !== 'undefined'
  //           && Array.prototype.slice.call(event.target.attributes, 0).filter(function (item) {
  //               return item.nodeValue === 'checkbox';
  //       }).length > 0) {
  //           // This is a workaround for checkboxes not being checked in case of stopping propagation
  //       } else if (Object.prototype.toString.call(event) == '[object PointerEvent]') {
  //           event.stopPropagation();
  //       }
  //   }
  // , true);

    angular.module('starter', ['ionic', 'starter.controllers', 'starter.services', 'ngMessages'])
      .run(function ($ionicPlatform) {
          $ionicPlatform.ready(function () {
              // Hide the accessory bar by default (remove this to show the accessory bar above the keyboard
              // for form inputs)
              if (window.cordova && window.cordova.plugins && window.cordova.plugins.Keyboard) {
                  cordova.plugins.Keyboard.hideKeyboardAccessoryBar(true);
              }
              if (window.StatusBar) {
                  // org.apache.cordova.statusbar required
                  StatusBar.styleDefault();
              }
          });
      })

    .value('tenantName', 'sampleDirectory2015.onmicrosoft.com')
    .value('authority', 'https://login.windows.net/common/')
    .value('resourceUrl', 'https://graph.windows.net/')
    .value('appId', '98ba0820-f7da-4411-87bb-598e0475536b')
    .value('redirectUrl', 'http://localhost:4400/services/aad/redirectTarget.html');

})();
