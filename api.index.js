const { ConfidentialClientApplication } = require('@azure/msal-node');

const msalConfig = {
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    authority: 'https://login.microsoftonline.com/common'
  }
};

const cca = new ConfidentialClientApplication(msalConfig);
const REDIRECT_URI = process.env.REDIRECT_URI;
const FRONTEND_URL = process.env.FRONTEND_URL;

module.exports = async (req, res) => {
  const url = req.url;

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  if (url.startsWith('/auth/login')) {
    try {
      const authUrl = await cca.getAuthCodeUrl({
        scopes: ['Mail.Read', 'Mail.Send', 'User.Read', 'offline_access'],
        redirectUri: REDIRECT_URI
      });
      res.writeHead(302, { Location: authUrl });
      res.end();
    } catch(e) {
      res.status(500).json({ error: e.message });
    }
    return;
  }

  if (url.startsWith('/auth/callback')) {
    const urlParams = new URL(req.url, `https://${req.headers.host}`);
    const code = urlParams.searchParams.get('code');
    if (!code) { res.status(400).json({ error: 'No code provided' }); return; }
    try {
      const result = await cca.acquireTokenByCode({
        code,
        scopes: ['Mail.Read', 'Mail.Send', 'User.Read', 'offline_access'],
        redirectUri: REDIRECT_URI
      });
      const token = result.accessToken;
      const email = result.account.username;
      res.writeHead(302, { Location: `${FRONTEND_URL}/index/index.html?auth=success&token=${token}&email=${encodeURIComponent(email)}` });
      res.end();
    } catch(e) {
      res.writeHead(302, { Location: `${FRONTEND_URL}/index/index.html?auth=error&msg=${encodeURIComponent(e.message)}` });
      res.end();
    }
    return;
  }

  if (url.startsWith('/api/emails')) {
    const authHeader = req.headers.authorization;
    if (!authHeader) { res.status(401).json({ error: 'No token' }); return; }
    const token = authHeader.replace('Bearer ', '');
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=20&$orderby=receivedDateTime desc&$select=id,subject,from,isRead,receivedDateTime,bodyPreview,body', {
        headers: { Authorization: `Bearer ${token}` }
      });
      const data = await response.json();
      if (data.error) { res.status(400).json({ error: data.error.message }); return; }
      res.status(200).json(data.value);
    } catch(e) {
      res.status(500).json({ error: e.message });
    }
    return;
  }

  if (url === '/api/health' || url === '/') {
    res.status(200).json({ status: 'Voxride backend running', version: '2.0' });
    return;
  }

  res.status(404).json({ error: 'Route not found', url });
};
