export default async function handler(req, res) {
  const gasUrl = process.env.GAS_WEB_APP_URL;
  if (!gasUrl) {
    return res.status(500).json({ ok: false, error: 'Missing GAS_WEB_APP_URL environment variable.' });
  }

  try {
    if (req.method === 'GET') {
      const action = String(req.query.action || '');
      if (action !== 'getSharedData') {
        return res.status(400).json({ ok: false, error: 'Unsupported API action.' });
      }
      const upstream = await fetch(`${gasUrl}${gasUrl.includes('?') ? '&' : '?'}action=getSharedData`, {
        method: 'GET',
        redirect: 'follow',
        headers: { Accept: 'application/json' }
      });
      const data = await upstream.json();
      if (!upstream.ok) return res.status(502).json({ ok: false, error: 'Google Apps Script request failed.' });
      return res.status(200).json(data);
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
      if (!body || !['createRequest', 'recordContribution'].includes(body.action)) {
        return res.status(400).json({ ok: false, error: 'Unsupported API action.' });
      }
      const upstream = await fetch(gasUrl, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8', Accept: 'application/json' },
        body: JSON.stringify(body)
      });
      const data = await upstream.json();
      if (!upstream.ok) return res.status(502).json({ ok: false, error: 'Google Apps Script request failed.' });
      return res.status(200).json(data);
    }

    res.setHeader('Allow', 'GET, POST');
    return res.status(405).json({ ok: false, error: 'Method not allowed.' });
  } catch (error) {
    return res.status(502).json({ ok: false, error: error.message || 'Could not reach Google Apps Script.' });
  }
}
