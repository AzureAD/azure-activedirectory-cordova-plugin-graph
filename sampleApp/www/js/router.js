(function () {
  'use strict';

  angular.module('starter')
    .config(['$stateProvider', '$urlRouterProvider', route]);

  function route($stateProvider, $urlRouterProvider) {
    $stateProvider
      .state('app', {
        url: '/app',
        abstract: true,
        templateUrl: 'views/master.html',
        controller: 'AppCtrl'
      })
      .state('app.user-list', {
        url: '/users',
        views: {
          'pageContent': {
            templateUrl: 'views/user-list.html',
            controller: 'UserListCtrl as vm'
          }
        }
      })
      .state('app.user-detail', {
        url: '/users/:userId',
        views: {
          'pageContent': {
            templateUrl: 'views/user-detail.html',
            controller: 'UserDetailCtrl as vm'
          }
        }
      })
      .state('app.group-list', {
        url: '/groups',
        views: {
          'pageContent': {
            templateUrl: 'views/group-list.html',
            controller: 'GroupListCtrl as vm'
          }
        }
      })
      .state('app.group-detail', {
        url: '/groups/:groupId',
        views: {
          'pageContent': {
            templateUrl: 'views/group-detail.html',
            controller: 'GroupDetailCtrl as vm'
          }
        }
      })
      .state('app.application-list', {
          url: '/applications',
          views: {
              'pageContent': {
                  templateUrl: 'views/application-list.html',
                  controller: 'AppListCtrl as vm'
              }
          }
      })
      .state('app.application-detail', {
          url: '/applications/:objectId',
          views: {
              'pageContent': {
                  templateUrl: 'views/application-detail.html',
                  controller: 'AppDetailCtrl as vm'
              }
          }
      })
      .state('app.deleted-app-list', {
          url: '/deleted-apps',
          views: {
              'pageContent': {
                  templateUrl: 'views/deleted-app-list.html',
                  controller: 'DeletedAppListCtrl as vm'
              }
          }
      })
      .state('app.deleted-app-detail', {
          url: '/deleted-apps/:objectId',
          views: {
              'pageContent': {
                  templateUrl: 'views/deleted-app-detail.html',
                  controller: 'DeletedAppDetailCtrl as vm'
              }
          }
      });

    // if none of the above states are matched, use this as the fallback
    $urlRouterProvider.otherwise('/app/users');
  }

})();