/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

var express = require('express');
var router = express.Router();
const { Client } = require("@microsoft/microsoft-graph-client");

var fetch = require('../fetch');

var { GRAPH_ME_ENDPOINT } = require('../authConfig');

// custom middleware to check auth state
function isAuthenticated(req, res, next) {
    if (!req.session.isAuthenticated) {
        return res.redirect('/auth/signin'); // redirect to sign-in route
    }

    next();
};

router.get('/id',
    isAuthenticated, // check if user is authenticated
    async function (req, res, next) {
        res.render('id', { idTokenClaims: req.session.account.idTokenClaims });
    }
);

router.get('/profile',
    isAuthenticated, // check if user is authenticated
    async function (req, res, next) {
        try {
            const graphResponse = await fetch(GRAPH_ME_ENDPOINT, req.session.accessToken);
            res.render('profile', { profile: graphResponse });
        } catch (error) {
            next(error);
        }
    }
);

router.get('/mail',
    isAuthenticated,
    async function (req, res, next) {
        try {
            const userAccessToken = req.session.accessToken;
            const client = Client.init({
                authProvider: (done) => {
                  done(null, userAccessToken);
                },
            });

            const messages = await client.api('/me/messages').get();
            res.render('mail', { messages: messages });
        } catch (error) {
            console.log(error)
            next(error);
        }
    }
);

module.exports = router;
