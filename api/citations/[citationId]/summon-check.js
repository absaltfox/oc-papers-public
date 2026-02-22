import { getCitationForSummon } from '../../../src/db.js';

const SUMMON_API = 'https://ubc.summon.serialssolutions.com/api/search';
const SUMMON_WEB = 'https://ubc.summon.serialssolutions.com/#!/search';
const ILL_URL = 'https://ill-docdel.library.ubc.ca/home';

function buildQuery(row) {
  if (row.query_title) {
    let q = `Title:(${row.query_title})`;
    if (row.query_author) q += ` AND Author:(${row.query_author})`;
    return q;
  }
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
    const apiUrl = `${SUMMON_API}?pn=1&l=en&include.ft.matches=t&q=${encodeURIComponent(q)}`;
    const resp = await fetch(apiUrl, {
      headers: { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' },
      signal: AbortSignal.timeout(8000),
    });
    if (!resp.ok) throw new Error(`Summon responded ${resp.status}`);
    const data = await resp.json();

    const results = (data.documents || []).slice(0, 10).map((d) => ({
      title: String(d.title || '').replace(/<\/?mark>/g, ''),
      authors: (d.authors || []).map((a) => a.fullname || a.name || '').filter(Boolean).join(', '),
      contentType: d.content_type || '',
      year: d.publication_date || '',
      inHoldings: d.in_holdings === true,
      link: d.link || '',
      snippet: String(d.snippet || '').replace(/<\/?mark>/g, ''),
    }));

    const found = results.some((r) => r.inHoldings);
    const searchUrl = `${SUMMON_WEB}?q=${encodeURIComponent(q)}`;

    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ found, results, searchUrl, illUrl: ILL_URL }));
  } catch (err) {
    res.statusCode = 502;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: 'Summon lookup failed' }));
  }
}
