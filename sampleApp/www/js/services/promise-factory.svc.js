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