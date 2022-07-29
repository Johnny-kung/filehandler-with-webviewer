const express = require('express');
const app = express();
const port = process.env.PORT || 3000;
const axios = require('axios');
const authFactory = require('./helpers/authFactory.helper') ;
const { toBase64, fromBase64 } = require('./utils/index');
const applyMiddleWares = require("./applyMiddleWares");
require('dotenv').config();

applyMiddleWares(app);

app.post('/webviewer', async (req, res) => {
    const { token } = req.session;
    console.log(token);
    if (token) {
        const activationItems = `${JSON.parse(req.body.items)[0]}`;
        req.session.activationItems = activationItems;
        const resp = await axios.get(activationItems, {
            headers: {
              'authorization': `Bearer ${token}`
            },
        });
        
        const { "@microsoft.graph.downloadUrl": downloadUrl } = resp.data
        res.redirect(`/index.html?filepath=${downloadUrl}`);
    } else {
        const state = toBase64(JSON.stringify({
            target: req.url, 
            activationParams: {
                items: req.body.items
            }
        }));

        const tenantId = process.env.TENANT_ID;
        const authClient = authFactory({
            auth: {
                clientId: process.env.CLIENT_ID,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                clientSecret: process.env.CLIENT_SECRET
            }
        });
        const authUrl = await authClient.getAuthCodeUrl({
            scopes: ["openid", "Files.ReadWrite.Selected"],
            redirectUri: 'http://localhost:3000/api/auth/login',
            state
        });
        console.log(authUrl);
    
        res.set('x-frame-options', 'SAMEORIGIN');
        res.writeHead(302, {
            location: authUrl,
        });
        res.end();
    }
});

app.get('/webviewer', async (req, res) => {
    const loginState = JSON.parse(fromBase64(req.query.state));
    const resp = await axios.get(`${JSON.parse(loginState.activationParams.items)[0]}`, {
        headers: {
          'authorization': `Bearer ${req.query.token}`
        },
    });
    req.session.activationItems = `${JSON.parse(loginState.activationParams.items)[0]}`;
    const { "@microsoft.graph.downloadUrl": downloadUrl } = resp.data
    res.redirect(`/index.html?filepath=${downloadUrl}`);
});

app.get('/api/auth/login', async (req, res) => {
    const tenantId = process.env.TENANT_ID;
    const authClient = authFactory({
        auth: {
            clientId: process.env.CLIENT_ID,
            authority: `https://login.microsoftonline.com/${tenantId}`,
            clientSecret: process.env.CLIENT_SECRET
        }
    });

    const state = JSON.parse(fromBase64(req.query.state));
    console.log('target url', state.target);
    const { code } = req.query;

    const tokenResp = await authClient.acquireTokenByCode({
        code,
        redirectUri: 'http://localhost:3000/api/auth/login',
        scopes: ['openid', 'Files.ReadWrite.All']
    });
    console.log('set session token', tokenResp.accessToken);
    req.session.token = tokenResp.accessToken;

    res.redirect(`${state.target}?state=${req.query.state}&token=${tokenResp.accessToken}&expiresOn=${tokenResp.expiresOn}`);
});

app.get('/token', async (req, res) => {
    console.log(req.session.token);
    if (req.session.token) {
        res.send({
            message: 'success',
            data: {
                token: req.session.token,
                activationItems: req.session.activationItems
            }
        })
    } else {
        res.send({
            status: 'fail',
            message: "Couldn't get token."
        });
    }
});

app.listen(port, () => {
    console.log('example app listening on ', port);
});
