---
page_type: sample
products:
- office-365
languages:
- javascript
title: Microsoft Teams NodeJS Helloworld Tabs SSO Sample
description: Microsoft Teams hello world sample app for tabs SSO in Node.js
extensions:
  contentType: samples
  createdDate: 11/3/2017 12:53:17 PM
---

# Authentication documentation

## Prerequisites

1. Setup an [Ngrok](https://ngrok.com/) account. This will allow you to test locally.
    * You will also need to sign up for a reserved subdomain since you will need a permanent address to use as your app's URL in Azure AD.
    * Make sure you've downloaded and installed Ngrok on your local machine. 
2. Create an [AAD application](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#1-create-your-aad-application-in-azure) in Azure. You can do this by visiting the "Azure AD app registration" portal in Azure.
    * Set your application URI to the same URI you've created in Ngrok. 
        * Ex: `api://contoso.ngrok.io`
    * Setup your redirect URIs. This will allow Azure AD to return authentication results to the correct URI.
        * Visit `Manage > Authentication`. 
        * Create a redirect URI in the format of: `https://contoso.ngrok.io/auth/auth-end`.
        * Enable Implicit Grant by selecting `Access Tokens` and `ID Tokens`.
    * Setup a client secret. You will need this when you exchange the token for more API permissions from your backend.
        * Visit `Manage > Certificates & secrets`
        * Create a new client secret.
    * Setup your API permissions. This is what your application is allowed to request permission to access.
        * Visit `Manage > API Permissions`
        * Make sure you have the following Graph permissions enabled: `email`, `offline_access`, `openid`, `User.Read` and `profile`.
        * Our SSO flow will give you access to the first 4 permissions, and we will have to exchange the token server-side to get an elevated token for the `profile` permission.
    * Expose an API that will give the Teams desktop, web and mobile clients access to the permissions above
        * Visit `Manage > Expose an API`
        * Add a scope and give it a scope name of `access_as_user`. Your API url should look like this: `api://contoso.ngrok.io/{appID}/access_as_user`. In the "who can consent" step, enable it for "Admins and users". Make sure the state is set to "enabled".
        * Next, add two client applications. This is for the Teams desktop/mobile clients and the web client.
            * 5e3ce6c0-2b1f-4285-8d4b-75ee78787346
            * 1fec8e78-bce4-4aaf-ab1b-5451cc387264

## Configuring this app for local development

1. Update your `manifest.json` file
    * Replace `{ngrokSubdomain}` with the subdomain you've assigned to your Ngrok account in step #1 above.
    * Update your `webApplicationInfo` section with your Azure AD application ID that you were assigned in step #2 above.
2. Update your `config/default.json` file
    * Replace the `tab.id` property with you Azure AD application ID
    * Replace the `tab.password` property with the "client secret" you were assigned in step #2

## Running the app locally

1. Run Ngrok to expose your local web server via a public URL. Make sure to point it to your Ngrok URI
    * Win: `./ngrok http 3333 -host-header=localhost:3333 -subdomain="contoso"`
    * Mac: `/ngrok http 3333 -host-header=localhost:3333 -subdomain="contoso"`
2. Install the neccessary NPM packages and start the app
    * `npm install`
    * `npm start`
    * Your app should be running on port 3333

## Packaging and installing your app to Teams

1. Package your manifest 
    * `gulp package-manifest`
    * This will create a zip file in the manifest folder
2. Install in Teams
    * Open Teams and visit the app store. Depending on the version of Teams, you may see an "App Store" button in the bottom left of Teams or you can find the app store by visiting `Apps > More Apps` in the left-hand app rail.
    * Install the app by clicking on the `Upload a custom app` link in the bottom left-hand side of the app store.

## Trying out the app

1. Once you've installed the app, it should automatically open for you. Visit the `Auth Tab` to begin testing out the authentication flow.
2. Follow the onscreen prompts. The authentication flow will print the output to your screen.
    * First you will see a dialog in Teams letting you know you need to grant permission to use the app.
    * Then an AAD consent dialog will appear in a seperate window.
    * Once you've consented to the permissions, your app will have enough information to authenticate you. This should be enough for _most_ apps.
    * The app will then try and access an API server-side for which you have not yet granted permission but fail in doing so. You will then see the "Grant further permission" button enabled. You can click this to trigger another consent dialog for the additional permissions.
    * Once you've granted all the permissions, you can revisit this tab and you will notice that you will automatically be logged in.

## Additional reading

More information for this sample - and for how to get started with Microsoft Teams development in general - is found in [Get started on the Microsoft Teams platform with Node.js and App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio).

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

