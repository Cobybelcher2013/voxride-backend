module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const authHeader = req.headers.authorization;
  if (!authHeader) { res.status(401).json({ error: 'No token provided' }); return; }
  
  const token = authHeader.replace('Bearer ', '');
  try {
    const response = await fetch(
      'https://graph.microsoft.com/v1.0/me/messages?$top=20&$orderby=receivedDateTime desc&$select=id,subject,from,isRead,receivedDateTime,bodyPreview,body',
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = await response.json();
    if (data.error) { res.status(400).json({ error: data.error.message }); return; }
    res.status(200).json(data.value);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
};
