
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function() {
    'use strict';

    angular
        .module('starter.services')
        .factory('PromiseFactory', [PromiseFactory]);

    function PromiseFactory() {
        var Deferred = cordova.require('cordova-plugin-ms-adal.utility').Utility.Deferred;

        var factory = { createPromise: createPromise };
        return factory;

        function createPromise() {
            return new Deferred();
        }
    }

}());