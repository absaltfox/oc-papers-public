import { createClient } from '@libsql/client/web';
import path from 'node:path';

const url = (process.env.TURSO_DATABASE_URL || `file:${path.join(process.cwd(), 'data', 'metrics.sqlite')}`).trim();
const authToken = process.env.TURSO_AUTH_TOKEN;

const client = createClient({ url, authToken });

// --- Document functions ---

export async function listAllDocumentMetadata() {
  const result = await client.execute('SELECT doc_id, metadata_json FROM documents');
  return result.rows.map((row) => {
    try {
      return { docId: row.doc_id, metadata: JSON.parse(row.metadata_json) };
    } catch {
      return null;
    }
  }).filter(Boolean);
}

// --- File metrics ---

export async function listAllFileMetrics() {
  const result = await client.execute(`
    SELECT doc_id, download_url, file_bytes, word_count, page_count,
           word_source, page_source, status
    FROM file_metrics
  `);
  const map = new Map();
  for (const row of result.rows) {
    map.set(row.doc_id, row);
  }
  return map;
}

// --- Committee members ---

export async function listAllCommitteeMembers() {
  const result = await client.execute(`
    SELECT doc_id, name, role, affiliation
    FROM committee_members
    ORDER BY id
  `);
  const map = new Map();
  for (const row of result.rows) {
    if (!map.has(row.doc_id)) map.set(row.doc_id, []);
    map.get(row.doc_id).push(row);
  }
  return map;
}

// --- Citation functions ---

export async function loadDocumentCitationsWithSharing(docId) {
  const result = await client.execute({
    sql: `
      SELECT c.id, c.citation_hash, c.citation_text,
        (SELECT COUNT(*) FROM document_citations dc2 WHERE dc2.citation_id = c.id) as total_docs,
        cl.hits AS catalogue_hits,
        cl.query_author AS catalogue_query_author,
        cl.query_title AS catalogue_query_title,
        cl.bib_id AS catalogue_bib_id,
        cl.looked_up_at AS catalogue_looked_up_at
      FROM citations c
      JOIN document_citations dc ON dc.citation_id = c.id
      LEFT JOIN catalogue_lookups cl ON cl.citation_id = c.id
      WHERE dc.doc_id = ?
      ORDER BY total_docs DESC, c.citation_text
    `,
    args: [docId]
  });
  return result.rows;
}

export async function loadDocsByCitation(citationId) {
  const result = await client.execute({
    sql: `
      SELECT d.doc_id as id, json_extract(d.metadata_json, '$.title') as title,
        json_extract(d.metadata_json, '$.author') as author
      FROM document_citations dc
      JOIN documents d ON d.doc_id = dc.doc_id
      WHERE dc.citation_id = ?
      ORDER BY title
    `,
    args: [citationId]
  });
  return result.rows;
}

export async function getTopCitedWorks(limit = 50) {
  const result = await client.execute({
    sql: `
      SELECT c.id, c.citation_text,
        COUNT(DISTINCT dc.doc_id) AS doc_count,
        cl.hits AS catalogue_hits,
        cl.bib_id AS catalogue_bib_id
      FROM citations c
      JOIN document_citations dc ON dc.citation_id = c.id
      LEFT JOIN catalogue_lookups cl ON cl.citation_id = c.id
      GROUP BY c.id
      HAVING doc_count > 1
      ORDER BY doc_count DESC, c.citation_text
      LIMIT ?
    `,
    args: [limit]
  });
  return result.rows;
}
