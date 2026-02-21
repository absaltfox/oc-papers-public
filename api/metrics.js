import { collectMetricsFromDb } from '../src/metrics.js';

let cache = null;
let cacheTs = 0;
const TTL = 10 * 60 * 1000;

export default async function handler(req, res) {
  if (Date.now() - cacheTs < TTL && cache) {
    res.setHeader('content-type', 'application/json');
    res.end(cache);
    return;
  }
  try {
    const payload = await collectMetricsFromDb();
    cache = JSON.stringify(payload);
    cacheTs = Date.now();
    res.setHeader('content-type', 'application/json');
    res.end(cache);
  } catch (err) {
    res.statusCode = 500;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ error: err?.message, stack: err?.stack }));
  }
}
