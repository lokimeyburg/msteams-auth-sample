(function() {
    'use strict';

    function printLog(msg) {
        var logDiv = document.getElementById('logs');
        var p = document.createElement("p");
        logDiv.prepend(msg, p);
    }

    // Start authentication
    // var states = ["getAuthToken", "checkPermissions", "needMorePermissions", "permissionsGranted"]
    function getAuthToken(){
        // Get auth token
        var authTokenRequest = {
            successCallback: function(result) { console.log("Success: " + result); },
            failureCallback: function(error) { console.log("Failure: " + error); },
        };

        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }


    printLog("Starting...");
    printLog("Getting auth token...");

})();
