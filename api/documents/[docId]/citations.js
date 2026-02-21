import { loadDocumentCitationsWithSharing } from '../../../src/db.js';

export default async function handler(req, res) {
  const { docId } = req.query;
  if (!docId || !/^\d+\.\d+$/.test(docId)) {
    res.statusCode = 400;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Invalid docId' }));
    return;
  }
  try {
    const citations = await loadDocumentCitationsWithSharing(docId);
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ citations }));
  } catch (err) {
    res.statusCode = 500;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Failed to load citations' }));
  }
}
