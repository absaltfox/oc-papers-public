export default function handler(req, res) {
  res.setHeader('content-type', 'application/json');
  res.end(JSON.stringify({ ok: true, timestamp: new Date().toISOString() }));
}
