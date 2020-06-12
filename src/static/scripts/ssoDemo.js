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

        return new Promise((resolve, reject) => {

            printLog("1. Get auth token from Microsoft Teams");

            microsoftTeams.authentication.getAuthToken({
                successCallback: (result) => {
                    printLog(result)
                    resolve(result);
                },
                failureCallback: function (error) {
                    reject("Error getting token: " + error);
                }
            });

        });

    }

    // 2. Exchange that token for a token with the required permissions
    //    using the web service (see /auth/token handler in app.js)
    function step2_ExchangeForServerSideToken(clientSideToken) {

        printLog("2. Exchange for server-side token");

        return new Promise((resolve, reject) => {

            microsoftTeams.getContext((context) => {

                fetch('/auth/token', {
                    method: 'post',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ "tid": context.tid, "token": clientSideToken }),
                    mode: 'cors',
                    cache: 'default'
                })
                .then((response) => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        reject(response.error);
                    }
                })
                .then((responseJson) => {
                    if (responseJson.error) {
                        reject(responseJson.error);
                    } else {
                        const serverSideToken = responseJson;
                        printLog(serverSideToken);
                        resolve(serverSideToken);
                    }
                });
            });
        });
    }

    // 3. Get the server side token and use it to call the Graph API
    function step3_UseServerSideToken(data) {

        return fetch("https://graph.microsoft.com/v1.0/me/",
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
    step1_GetAuthToken()
        .then((clientSideToken) => {
            return step2_ExchangeForServerSideToken(clientSideToken);
        })
        .then((serverSideToken) => {
            return step3_UseServerSideToken(serverSideToken);
        })
        .catch((error) => {
            if (error === "invalid_grant") {
                printLog("Server needs user consent - enable the 'Grant Permission' button");
                document.getElementById("promptForConsentButton").disabled = false;
            } else {
                printLog(`Error from web service: ${error}`);
            }
        });

})();
