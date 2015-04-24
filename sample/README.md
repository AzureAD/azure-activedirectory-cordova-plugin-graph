### How to setup the sample app: ###

* Start with `git clone https://github.com/AzureAD/azure-activedirectory-cordova-plugin-graph`
* `cordova create sample --copy-from="azure-activedirectory-cordova-plugin-graph/sample/www"`
* `cd sample`
* `cordova plugin add cordova-plugin-device`
* `cordova plugin add cordova-plugin-console`
* `cordova plugin add com.ionic.keyboard`
* `cordova plugin add ../azure-activedirectory-cordova-plugin-graph`
* `cordova platforms add android`
* `cordova platforms add windows`
* `cordova run android`
* `cordova run windows`
