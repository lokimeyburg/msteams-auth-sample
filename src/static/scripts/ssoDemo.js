(function () {
    'use strict';

    // Set up button to get additional consent
    // This button will be enabled if the on-behalf-of-flow fails
    // due to requiring further consent. It shows the user an AAD
    // consent dialog and asks for additional permission
    function initializeConsentButton() {
        var btn = document.getElementById("promptForConsentButton")
        btn.onclick = () => {
            microsoftTeams.authentication.authenticate({
                url: window.location.origin + "/auth/auth-start",
                width: 600,
                height: 535,
                successCallback: function (result) {
                    let data = localStorage.getItem(result);
                    localStorage.removeItem(result);
                    printLog("Success! Additional permission granted. Result: " + data)
                    window.location.reload();
                },
                failureCallback: function (reason) {
                    printLog("Failure. Additional permission was not granted. Result: " + JSON.stringify(reason))
                }
            });

        }
    }

    // 1. Get auth token
    // Ask Teams to get us a token from AAD
    function step1_GetAuthToken() {

        printLog("1. Get auth token from Microsoft Teams");

        // Get auth token
        var authTokenRequest = {
            successCallback: (result) => {
                printLog(result)
                step2_ExchangeForServerSideToken(result);
            },
            failureCallback: function (error) {
                printLog("Error getting token: " + error);
            },
        };

        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    // 2. Exchange that token for a token with the required permissions
    //    using the web service (see /auth/token handler in app.js)
    function step2_ExchangeForServerSideToken(result) {
        printLog("2. Exchange for server-side token")

        // Get Tenant ID
        var getContextPromise = new Promise((resolve, reject) => {
            microsoftTeams.getContext(function (context) {
                resolve(context);
            });
        });

        // Send Tenant ID and token to backend
        getContextPromise.then(function (context) {
            // POST result to backend
            var xhr = new XMLHttpRequest();
            xhr.open("POST", "/auth/token", true);
            xhr.setRequestHeader('Content-Type', 'application/json');
            // Set response handler before sending
            xhr.onreadystatechange = function () {
                if (this.readyState != 4) return;
                if (this.status == 200) {
                    var data = JSON.parse(this.responseText);
                    printLog(data);
                    step3_UseServerSideToken(data);
                }
            };
            // send POST request
            xhr.send(JSON.stringify({ "tid": context.tid, "token": result }));
        });
    }

    // 3. Get the server side token and use it to call the Graph API
    function step3_UseServerSideToken(data) {

        var error = data.error;
        if (error != null) {
            // Error: enable the grantPermission button
            printLog("Server needs user consent - enable the 'Grant Permission' button");
            document.getElementById("promptForConsentButton").disabled = false;
        } else {
            // Success: server returned a valid acess token
            printLog("3. Use token to get user profile from Graph API");
            fetch("https://graph.microsoft.com/v1.0/me/",
                {
                    method: 'GET',
                    headers: {
                        "accept": "application/json",
                        "authorization": "bearer " + data
                    },
                    mode: 'cors',
                    cache: 'default'
                })
                .then((response) => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        throw (`Error ${response.status}: ${response.statusText}`);
                    }
                })
                .then((profile) => {
                    printLog(JSON.stringify(profile, undefined, 4), 'pre');
                });
        }

    }

    // Add text to the display in a <p> or other HTML element
    function printLog(text, elementTag) {
        var logDiv = document.getElementById('logs');
        var p = document.createElement(elementTag ? elementTag : "p");
        p.innerText = text;
        logDiv.append(p);
        console.log("ssoDemo: " + text);
    }

    // In-line code
    initializeConsentButton();
    step1_GetAuthToken();

})();
