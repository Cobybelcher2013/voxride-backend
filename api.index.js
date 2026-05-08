const express = require('express');
const session = require('express-session');
const cors = require('cors');
const axios = require('axios');
const msal = require('@azure/msal-node');

const app = express();

app.use(cors({
  origin: process.env.FRONTEND_URL,
  credentials: true
}));

app.use(express.json());
app.use(session({
  secret: 'voxride-secret-key',
  resave: false,
  saveUninitialized: false
}));

const msalConfig = {
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/common`
  }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

// Login route
app.get('/auth/login', async (req, res) => {
  const authUrl = await cca.getAuthCodeUrl({
    scopes: ['Mail.Read', 'Mail.Send', 'User.Read'],
    redirectUri: process.env.REDIRECT_URI
  });
  res.redirect(authUrl);
});

// Callback route
app.get('/auth/callback', async (req, res) => {
  try {
    const result = await cca.acquireTokenByCode({
      code: req.query.code,
      scopes: ['Mail.Read', 'Mail.Send', 'User.Read'],
      redirectUri: process.env.REDIRECT_URI
    });
    req.session.accessToken = result.accessToken;
    req.session.userEmail = result.account.username;
    res.redirect(`${process.env.FRONTEND_URL}/index/index.html?auth=success`);
  } catch (err) {
    res.redirect(`${process.env.FRONTEND_URL}/index/index.html?auth=error`);
  }
});

// Get emails route
app.get('/api/emails', async (req, res) => {
  if (!req.session.accessToken) {
    return res.status(401).json({ error: 'Not authenticated' });
  }
  try {
    const response = await axios.get('https://graph.microsoft.com/v1.0/me/messages?$top=20&$orderby=receivedDateTime desc', {
      headers: { Authorization: `Bearer ${req.session.accessToken}` }
    });
    res.json(response.data.value);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get user info
app.get('/api/user', async (req, res) => {
  if (!req.session.accessToken) {
    return res.status(401).json({ error: 'Not authenticated' });
  }
  try {
    const response = await axios.get('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${req.session.accessToken}` }
    });
    res.json(response.data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Check auth status
app.get('/api/auth-status', (req, res) => {
  res.json({
    authenticated: !!req.session.accessToken,
    email: req.session.userEmail || null
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Voxride backend running on port ${PORT}`));

module.exports = app;
