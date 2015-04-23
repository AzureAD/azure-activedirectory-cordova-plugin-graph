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