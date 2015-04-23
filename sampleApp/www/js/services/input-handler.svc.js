(function () {
    'use strict';

    angular
        .module('starter.services')
        .factory('InputHandler', InputHandler);

    function InputHandler() {
        return {
            uriStringToArray: uriStringToArray
        };

        function uriStringToArray(urisString) {
            var uris = [];

            if (urisString.indexOf(',') !== -1) {
                uris = urisString.split(/\s*,\s*/);
            } else if (urisString.indexOf(';') !== -1) {
                uris = urisString.split(/\s*;\s*/);
            } else if (urisString.indexOf(' ') !== -1) {
                uris = urisString.split(/\s+/);
            } else {
                uris[0] = urisString;
            }

            for (var i = 0; i < uris.length; i++) {
                uris[i] = uris[i].trim();
            }

            return uris;
        }
    }

})();