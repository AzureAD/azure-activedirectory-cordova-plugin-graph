var bootstraper = {
  initialize: function () {
    document.addEventListener('deviceready', this.onDeviceReady, false);
  },
  onDeviceReady: function () {
    angular.bootstrap(document, ['starter']);
  }
};