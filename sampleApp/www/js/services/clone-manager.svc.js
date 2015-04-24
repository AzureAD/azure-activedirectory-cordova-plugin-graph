
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

(function() {
    'use strict';

    angular
        .module('starter.services')
        .factory('CloneManager', CloneManager);

    function CloneManager() {
        return {
            clone: clone
        };

        function clone(obj) {
            return createClone.apply(obj);

            function createClone() {
                var func = function () { };
                func.prototype = this;
                return new func();
            }
        }
    }
   
})();