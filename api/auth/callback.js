const { ConfidentialClientApplication } = require('@azure/msal-node');

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    authority: 'https://login.microsoftonline.com/common'
  }
});

module.exports = async (req, res) => {
  const code = req.query.code;
  if (!code) { res.status(400).send('No code provided'); return; }
  try {
    const result = await cca.acquireTokenByCode({
      code,
      scopes: ['Mail.Read', 'Mail.Send', 'User.Read', 'offline_access'],
      redirectUri: process.env.REDIRECT_URI
    });
    const token = result.accessToken;
    const email = result.account.username;
    const frontendUrl = process.env.FRONTEND_URL;
    res.redirect(`${frontendUrl}/index/index.html?auth=success&token=${encodeURIComponent(token)}&email=${encodeURIComponent(email)}`);
  } catch(e) {
    res.redirect(`${process.env.FRONTEND_URL}/index/index.html?auth=error`);
  }
};
