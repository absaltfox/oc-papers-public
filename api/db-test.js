import { createClient } from '@libsql/client/web';

export default async function handler(req, res) {
  const url = process.env.TURSO_DATABASE_URL;
  const authToken = process.env.TURSO_AUTH_TOKEN;
  try {
    const client = createClient({ url, authToken });
    const result = await client.execute('SELECT COUNT(*) as count FROM documents');
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ ok: true, url, count: result.rows[0].count }));
  } catch (err) {
    res.statusCode = 500;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ ok: false, url, error: err?.message, stack: err?.stack }));
  }
}
