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
    SELECT doc_id, name, role, affiliation, source
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

// --- Summon lookup helper ---

export async function getCitationForSummon(citationId) {
  const result = await client.execute({
    sql: `SELECT c.citation_text, cl.query_title, cl.query_author
          FROM citations c
          LEFT JOIN catalogue_lookups cl ON cl.citation_id = c.id
          WHERE c.id = ?`,
    args: [citationId]
  });
  return result.rows[0] || null;
}

// --- Citation counts per document ---

export async function listAllCitationCounts() {
  const result = await client.execute(`
    SELECT doc_id, COUNT(*) as count
    FROM document_citations
    GROUP BY doc_id
  `);
  const map = new Map();
  for (const row of result.rows) {
    map.set(row.doc_id, Number(row.count));
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

// --- Topic functions ---

export async function hasTopics() {
  try {
    const result = await client.execute('SELECT 1 FROM topics LIMIT 1');
    return result.rows.length > 0;
  } catch {
    return false;
  }
}

export async function loadTopics() {
  const result = await client.execute('SELECT topic_id, label, top_terms, doc_count, model_name, created_at FROM topics ORDER BY doc_count DESC');
  return result.rows.map((row) => ({
    topicId: Number(row.topic_id),
    label: row.label,
    topTerms: (() => { try { return JSON.parse(row.top_terms); } catch { return []; } })(),
    docCount: Number(row.doc_count),
    modelName: row.model_name,
    createdAt: row.created_at,
  }));
}

export async function loadDocumentTopics() {
  const result = await client.execute('SELECT doc_id, topic_id, probability FROM document_topics');
  const map = new Map();
  for (const row of result.rows) {
    map.set(row.doc_id, { topicId: Number(row.topic_id), probability: row.probability != null ? Number(row.probability) : null });
  }
  return map;
}

export async function loadDocumentTopicCoords() {
  try {
    const result = await client.execute('SELECT doc_id, umap_x, umap_y FROM document_topic_coords');
    const map = new Map();
    for (const row of result.rows) {
      map.set(row.doc_id, { x: Number(row.umap_x), y: Number(row.umap_y) });
    }
    return map;
  } catch {
    return new Map();
  }
}

export async function loadTopicHierarchy() {
  try {
    const result = await client.execute('SELECT leaf_topic_ids, linkage_json FROM topic_hierarchy_meta WHERE id = 1');
    if (!result.rows.length) return null;
    const row = result.rows[0];
    return {
      leafTopicIds: JSON.parse(row.leaf_topic_ids),
      linkage: JSON.parse(row.linkage_json),
    };
  } catch { return null; }
}

export async function getCitationCooccurrence(limit = 100) {
  const result = await client.execute({
    sql: `
      WITH top_citations AS (
        SELECT citation_id, COUNT(DISTINCT doc_id) AS cnt
        FROM document_citations
        GROUP BY citation_id
        HAVING cnt >= 2
        ORDER BY cnt DESC
        LIMIT 50
      )
      SELECT
        c1.id AS id1, substr(c1.citation_text, 1, 80) AS text1, tc1.cnt AS freq1,
        c2.id AS id2, substr(c2.citation_text, 1, 80) AS text2, tc2.cnt AS freq2,
        COUNT(DISTINCT dc1.doc_id) AS shared
      FROM document_citations dc1
      JOIN document_citations dc2
        ON dc1.doc_id = dc2.doc_id AND dc1.citation_id < dc2.citation_id
      JOIN citations c1 ON c1.id = dc1.citation_id
      JOIN citations c2 ON c2.id = dc2.citation_id
      JOIN top_citations tc1 ON tc1.citation_id = c1.id
      JOIN top_citations tc2 ON tc2.citation_id = c2.id
      GROUP BY dc1.citation_id, dc2.citation_id
      HAVING shared >= 2
      ORDER BY shared DESC
      LIMIT ?
    `,
    args: [limit]
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
