'use strict';

module.exports.setup = function(app) {
    var path = require('path');
    var express = require('express')
    
    // Configure the view engine, views folder and the statics path
    app.use(express.static(path.join(__dirname, 'static')));
    app.set('view engine', 'pug');
    app.set('views', path.join(__dirname, 'views'));
    
    // Setup home page
    app.get('/', function(req, res) {
        res.render('hello');
    });
    
    // Setup the static tab
    app.get('/hello', function(req, res) {
        res.render('hello');
    });
    
    // Setup the configure tab, with first and second as content tabs
    app.get('/configure', function(req, res) {
        res.render('configure');
    });    

    app.get('/first', function(req, res) {
        res.render('first');
    });
    
    app.get('/second', function(req, res) {
        res.render('second');
    }); 
    
    // Auth ---------------------------
    app.get('/auth', function(req, res) {
        res.render('auth');
    }); 

    // Token exchange
    app.post('/auth/token', function(req, res) {
        console.log(req);

        var tid = req.authInfo.tid
        var token = req.token
        var scopes = ["https://graph.microsoft.com/User.Read"]

        var oboPromise = new Promise((resolve, reject) => {
            const url = "https://login.microsoftonline.com/" + tid + "/oauth2/v2.0/token";
            const params = {
              client_id: process.env.APPSETTING_AAD_ApplicationId,
              client_secret: process.env.APPSETTING_AAD_ApplicationSecret,
              grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
              assertion: token,
              requested_token_use: "on_behalf_of",
              scope: scopes.join(" ")
            };
        
            fetch(url, {
              method: "POST",
              body: querystring.stringify(params),
              headers: {
                Accept: "application/json",
                "Content-Type": "application/x-www-form-urlencoded"
              }
            }).then(result => {
              if (result.status !== 200) {
                result.json().then(json => {
                  // TODO: Check explicitly for invalid_grant or interaction_required
                  reject(new ServerError(403, "ConsentRequired"));
                });
              } else {
                result.json().then(json => {
                  resolve(json.access_token);
                });
              }
            });
        });

        promise.then(function(result) {
            console.log(result); // "Stuff worked!"
            res.render(result);
        }, function(err) {
            console.log(err); // Error: "It broke"
            res.render(err);
        });
    });
};
