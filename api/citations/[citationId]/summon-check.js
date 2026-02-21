import { getCitationForSummon } from '../../../src/db.js';

const SUMMON_BASE = 'https://ubc.summon.serialssolutions.com/api/search';
const ILL_URL = 'https://ill-docdel.library.ubc.ca/home';

function buildQuery(row) {
  if (row.query_title) {
    let q = `Title:(${row.query_title})`;
    if (row.query_author) q += ` AND Author:(${row.query_author})`;
    return q;
  }
  // fallback: raw citation text, truncated
  return String(row.citation_text || '').slice(0, 200);
}

export default async function handler(req, res) {
  const { citationId } = req.query;
  if (!citationId || !/^\d{1,10}$/.test(citationId)) {
    res.statusCode = 400;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Invalid citationId' }));
    return;
  }

  const row = await getCitationForSummon(Number(citationId));
  if (!row) {
    res.statusCode = 404;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Citation not found' }));
    return;
  }

  const q = buildQuery(row);
  if (!q) {
    res.statusCode = 422;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Insufficient citation data to query Summon' }));
    return;
  }

  try {
    const url = `${SUMMON_BASE}?pn=1&l=en&include.ft.matches=t&q=${encodeURIComponent(q)}`;
    const resp = await fetch(url, {
      headers: { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' },
      signal: AbortSignal.timeout(8000),
    });
    if (!resp.ok) throw new Error(`Summon responded ${resp.status}`);
    const data = await resp.json();

    const held = (data.documents || []).find((d) => d.in_holdings === true);
    if (held) {
      const title = String(held.title || '').replace(/<\/?mark>/g, '');
      res.setHeader('content-type', 'application/json');
      res.end(JSON.stringify({ found: true, title, link: held.link || SUMMON_BASE }));
    } else {
      res.setHeader('content-type', 'application/json');
      res.end(JSON.stringify({ found: false, illUrl: ILL_URL }));
    }
  } catch (err) {
    res.statusCode = 502;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Summon lookup failed' }));
  }
}
