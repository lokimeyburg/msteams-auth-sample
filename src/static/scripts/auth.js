(function() {
    'use strict';

    // 1. Get auth token
    // Ask Teams to get us a token from AAD, we should exchange it when it n
    function getAuthToken(){
        // Get auth token
        var authTokenRequest = {
            successCallback: sendTokenToBackend(result),
            failureCallback: function(error) { console.log("Failure: " + error); },
        };

        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    // 2. Send token to the backend
    // After we call getAuthToken, we need to send the token from that request to the backend to 
    // verify that we have the correct permissions (or do an on-behalf-of exchange to get a new token) 
    function sendTokenToBackend(result){
        console.log(result)
        printLog("Token received: [TOKEN]")
        printLog("Sending token to backend for AAD on-behalf-of exchange")
        // POST result to backend
        var xhr = new XMLHttpRequest();
        xhr.open("POST", "/auth/token", true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        // Set response handler before sending
        xhr.onreadystatechange = function () {
            if (this.readyState != 4) return;
            if (this.status == 200) {
                var data = JSON.parse(this.responseText);
                console.log("we got the returned token from the backend");
            }
        };
        // send POST request
        xhr.send(JSON.stringify({
            value: value
        }));
    }

    // ------------------------------------------------------------------------
    
    function printLog(msg) {
        var logDiv = document.getElementById('logs');
        var p = document.createElement("p");
        logDiv.prepend(msg, p);
    }

    

    // Start authentication
    printLog("Starting...");
    printLog("Getting auth token...");
    getAuthToken();

})();
