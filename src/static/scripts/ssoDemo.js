(function () {
    'use strict';

    // Set up button to get additional consent
    // This button will be enabled if the on-behalf-of-flow fails
    // due to requiring further consent. It shows the user an AAD
    // consent dialog and asks for additional permission
    // function initializeConsentButton() {
    //     var btn = document.getElementById("promptForConsentButton")
    //     btn.onclick = () => {
    //         microsoftTeams.authentication.authenticate({
    //             url: window.location.origin + "/auth/auth-start",
    //             width: 600,
    //             height: 535,
    //             successCallback: function (result) {
    //                 let data = localStorage.getItem(result);
    //                 localStorage.removeItem(result);
    //                 printLog("Success! Additional permission granted. Result: " + data)
    //                 window.location.reload();
    //             },
    //             failureCallback: function (reason) {
    //                 printLog("Failure. Additional permission was not granted. Result: " + JSON.stringify(reason))
    //             }
    //         });
    //     }
    // }

    function requestConsent() {
        return new Promise((resolve, reject) => {
            microsoftTeams.authentication.authenticate({
                url: window.location.origin + "/auth/auth-start",
                width: 600,
                height: 535,
                successCallback: (result) => {
                    let data = localStorage.getItem(result);
                    localStorage.removeItem(result);
                    resolve(data);
                },
                failureCallback: (reason) => {
                    reject(JSON.stringify(reason));
                }
            });
        });
    }

    // 1. Get auth token
    // Ask Teams to get us a token from AAD
    function GetClientSideToken() {

        return new Promise((resolve, reject) => {

            display("1. Get auth token from Microsoft Teams");

            microsoftTeams.authentication.getAuthToken({
                successCallback: (result) => {
                    display(result)
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
    function GetServerSideToken(clientSideToken) {

        display("2. Exchange for server-side token");

        return new Promise((resolve, reject) => {

            microsoftTeams.getContext((context) => {

                fetch('/auth/token', {
                    method: 'post',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        'tid': context.tid,
                        'token': clientSideToken 
                    }),
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
                            display(serverSideToken);
                            resolve(serverSideToken);
                        }
                    });
            });
        });
    }

    // 3. Get the server side token and use it to call the Graph API
    function UseServerSideToken(data) {

        display("3. Call https://graph.microsoft.com/v1.0/me/ with the server side token");

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
                display(JSON.stringify(profile, undefined, 4), 'pre');
            });

    }

    // Add text to the display in a <p> or other HTML element
    function display(text, elementTag) {
        var logDiv = document.getElementById('logs');
        var p = document.createElement(elementTag ? elementTag : "p");
        p.innerText = text;
        logDiv.append(p);
        console.log("ssoDemo: " + text);
        return p;
    }

    // In-line code
    GetClientSideToken()
        .then((clientSideToken) => {
            return GetServerSideToken(clientSideToken);
        })
        .then((serverSideToken) => {
            return UseServerSideToken(serverSideToken);
        })
        .catch((error) => {
            if (error === "invalid_grant") {
                display(`Error: ${error} - user or admin consent required`);
                let button = display("Consent", "button");
                button.onclick = (() => {
                    requestConsent()
                        .then((result) => {
                            let accessToken = JSON.parse(result).accessToken;
                            display(`Received access token ${accessToken}`);
                            UseServerSideToken(accessToken);
                        })
                        .catch((error) => {
                            display(`ERROR ${error}`);
                            button.disabled = true;
                            let refreshButton = display("Refresh page", "button");
                            refreshButton.onclick = (() => { window.location.reload(); });
                        });
                });
            } else {
                display(`Error from web service: ${error}`);
            }
        });

})();
