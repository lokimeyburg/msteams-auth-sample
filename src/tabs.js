'use strict';
const fetch = require("node-fetch");
const querystring = require("querystring");

module.exports.setup = function(app) {
    var path = require('path');
    var express = require('express')
    
    // Configure the view engine, views folder and the statics path
    app.use(express.static(path.join(__dirname, 'static')));
    app.set('view engine', 'pug');
    app.set('views', path.join(__dirname, 'views'));
    // Use the JSON middleware
    app.use(express.json());
    
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
    
    // ------------------
    // Auth page
    app.get('/auth', function(req, res) {
        res.render('auth');
    }); 

    // Silent auth dialog
    app.get('/auth/silent-start', function(req, res) {
        var clientId = "bdb71ee3-1c28-4edb-a758-fd6f8b60348c"
        res.render('silent-start', { clientId: clientId });
    });

    // Silent auth end page
    app.get('/auth/silent-end', function(req, res) {
        var clientId = "bdb71ee3-1c28-4edb-a758-fd6f8b60348c"
        res.render('silent-end', { clientId: clientId });
    }); 

    // On-behalf-of token exchange
    app.post('/auth/token', function(req, res) {
        var tid = req.body.tid
        var token = req.body.token
        var scopes = ["https://graph.microsoft.com/User.Read"]

        var oboPromise = new Promise((resolve, reject) => {
            const url = "https://login.microsoftonline.com/" + tid + "/oauth2/v2.0/token";
            const params = {
                client_id: "bdb71ee3-1c28-4edb-a758-fd6f8b60348c",
                client_secret: "]DjvGB0f?R[Z4qSwn24uSfr?EKhGN_tv",
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
                  reject({"error":json.error});
                });
              } else {
                result.json().then(json => {
                  resolve(json.access_token);
                });
              }
            });
        });

        oboPromise.then(function(result) {
            console.log(result); // "Stuff worked!"
            res.json(result);
        }, function(err) {
            console.log(err); // Error: "It broke"
            res.json(err);
        });
    });
};
