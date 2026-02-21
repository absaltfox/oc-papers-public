import { getTopCitedWorks } from '../../src/db.js';
import { parseNumberParam } from '../../src/validate.js';

export default async function handler(req, res) {
  const limit = parseNumberParam(req.query?.limit, 50);
  try {
    const works = await getTopCitedWorks(Math.min(limit, 200));
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ works }));
  } catch (err) {
    res.statusCode = 500;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Failed to load top citations' }));
  }
}
