const { ConfidentialClientApplication } = require('@azure/msal-node');

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    authority: 'https://login.microsoftonline.com/common'
  }
});

module.exports = async (req, res) => {
  try {
    const authUrl = await cca.getAuthCodeUrl({
      scopes: ['Mail.Read', 'Mail.Send', 'User.Read', 'offline_access'],
      redirectUri: process.env.REDIRECT_URI
    });
    res.redirect(authUrl);
  } catch(e) {
    res.status(500).send('Auth error: ' + e.message);
  }
};
