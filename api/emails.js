module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  const token = req.headers.authorization?.replace('Bearer ', '');
  if (!token) { res.status(401).json({ error: 'No token' }); return; }
  try {
    const r = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=20&$orderby=receivedDateTime desc&$select=id,subject,from,isRead,receivedDateTime,bodyPreview,body', {
      headers: { Authorization: `Bearer ${token}` }
    });
    const data = await r.json();
    res.status(200).json(data.error ? { error: data.error.message } : data.value);
  } catch(e) { res.status(500).json({ error: e.message }); }
};
