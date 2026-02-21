import { loadDocsByCitation } from '../../../src/db.js';

export default async function handler(req, res) {
  const { citationId } = req.query;
  if (!citationId) {
    res.statusCode = 400;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Missing citationId' }));
    return;
  }
  try {
    const documents = await loadDocsByCitation(citationId);
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ documents }));
  } catch (err) {
    res.statusCode = 500;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Failed to load documents for citation' }));
  }
}
