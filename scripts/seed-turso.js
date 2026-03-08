/**
 * Seed Turso from local metrics.sqlite
 *
 * Usage:
 *   TURSO_DATABASE_URL=libsql://... TURSO_AUTH_TOKEN=... node scripts/seed-turso.js
 *
 * Reads from local metrics.sqlite using node:sqlite (DatabaseSync),
 * writes to Turso via @libsql/client in batches.
 */

import { DatabaseSync } from 'node:sqlite';
import { createClient } from '@libsql/client';
import path from 'node:path';

const SQLITE_PATH = process.env.SQLITE_PATH || path.join(process.cwd(), 'data', 'metrics.sqlite');
const TURSO_URL = process.env.TURSO_DATABASE_URL;
const TURSO_TOKEN = process.env.TURSO_AUTH_TOKEN;
const BATCH_SIZE = 200;

if (!TURSO_URL) {
  console.error('Error: TURSO_DATABASE_URL environment variable is required.');
  process.exit(1);
}

const src = new DatabaseSync(SQLITE_PATH);
const dest = createClient({ url: TURSO_URL, authToken: TURSO_TOKEN });

async function ensureSchema() {
  await dest.executeMultiple(`
    CREATE TABLE IF NOT EXISTS documents (
      doc_id TEXT PRIMARY KEY,
      metadata_json TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS file_metrics (
      doc_id TEXT PRIMARY KEY,
      pdf_path TEXT,
      download_url TEXT,
      file_bytes INTEGER,
      word_count INTEGER,
      page_count INTEGER,
      word_source TEXT,
      page_source TEXT,
      status TEXT,
      error TEXT,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS citations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      citation_hash TEXT UNIQUE NOT NULL,
      citation_text TEXT NOT NULL,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS document_citations (
      doc_id TEXT NOT NULL,
      citation_id INTEGER NOT NULL,
      updated_at TEXT NOT NULL,
      PRIMARY KEY (doc_id, citation_id),
      FOREIGN KEY (citation_id) REFERENCES citations(id)
    );

    CREATE TABLE IF NOT EXISTS catalogue_lookups (
      citation_id INTEGER PRIMARY KEY,
      hits INTEGER,
      query_author TEXT,
      query_title TEXT,
      bib_id TEXT,
      looked_up_at TEXT NOT NULL,
      FOREIGN KEY (citation_id) REFERENCES citations(id)
    );

    CREATE TABLE IF NOT EXISTS committee_members (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_id TEXT NOT NULL,
      name TEXT NOT NULL,
      role TEXT,
      affiliation TEXT,
      source TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      UNIQUE(doc_id, name, role)
    );

    CREATE TABLE IF NOT EXISTS topics (
      topic_id    INTEGER PRIMARY KEY,
      label       TEXT NOT NULL,
      top_terms   TEXT NOT NULL,
      doc_count   INTEGER NOT NULL,
      model_name  TEXT NOT NULL,
      created_at  TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS document_topics (
      doc_id      TEXT NOT NULL,
      topic_id    INTEGER NOT NULL,
      probability REAL,
      PRIMARY KEY (doc_id, topic_id)
    );

    CREATE TABLE IF NOT EXISTS document_topic_coords (
      doc_id  TEXT PRIMARY KEY,
      umap_x  REAL NOT NULL,
      umap_y  REAL NOT NULL
    );

    CREATE TABLE IF NOT EXISTS topic_hierarchy_meta (
      id              INTEGER PRIMARY KEY DEFAULT 1,
      leaf_topic_ids  TEXT NOT NULL,
      linkage_json    TEXT NOT NULL,
      created_at      TEXT NOT NULL
    );
  `);
  console.log('Schema ensured.');
}

async function seedTable(tableName, rows, buildStatement) {
  if (!rows.length) {
    console.log(`  ${tableName}: 0 rows, skipping.`);
    return;
  }
  let seeded = 0;
  for (let i = 0; i < rows.length; i += BATCH_SIZE) {
    const batch = rows.slice(i, i + BATCH_SIZE);
    const statements = batch.map((row) => buildStatement(row));
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        await dest.batch(statements, 'write');
        break;
      } catch (err) {
        if (attempt === 3) throw err;
        console.warn(`\n  ${tableName}: batch at ${seeded} failed (attempt ${attempt}), retrying in 5s...`);
        await new Promise((r) => setTimeout(r, 5000));
      }
    }
    seeded += batch.length;
    process.stdout.write(`\r  ${tableName}: ${seeded}/${rows.length}`);
    // Small delay every 2000 rows to avoid connection exhaustion
    if (seeded % 2000 === 0) await new Promise((r) => setTimeout(r, 500));
  }
  console.log(`\r  ${tableName}: ${seeded} rows seeded.`);
}

async function main() {
  console.log(`Source: ${SQLITE_PATH}`);
  console.log(`Destination: ${TURSO_URL}`);
  console.log('');

  await ensureSchema();

  // Seed documents
  const docs = src.prepare('SELECT doc_id, metadata_json, updated_at FROM documents').all();
  await seedTable('documents', docs, (row) => ({
    sql: `INSERT OR REPLACE INTO documents (doc_id, metadata_json, updated_at) VALUES (?, ?, ?)`,
    args: [row.doc_id, row.metadata_json, row.updated_at]
  }));

  // Seed file_metrics
  const fileMetrics = src.prepare(`
    SELECT doc_id, pdf_path, download_url, file_bytes, word_count, page_count,
           word_source, page_source, status, error, updated_at
    FROM file_metrics
  `).all();
  await seedTable('file_metrics', fileMetrics, (row) => ({
    sql: `INSERT OR REPLACE INTO file_metrics
      (doc_id, pdf_path, download_url, file_bytes, word_count, page_count, word_source, page_source, status, error, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    args: [
      row.doc_id, row.pdf_path ?? null, row.download_url ?? null,
      row.file_bytes ?? null, row.word_count ?? null, row.page_count ?? null,
      row.word_source ?? null, row.page_source ?? null, row.status ?? null,
      row.error ?? null, row.updated_at
    ]
  }));

  // Clear citation-related tables first to avoid FK constraint failures
  // (INSERT OR REPLACE on citations deletes+re-inserts, breaking FK refs)
  console.log('  Clearing citation tables for clean re-seed...');
  await dest.executeMultiple(`
    DELETE FROM catalogue_lookups;
    DELETE FROM document_citations;
    DELETE FROM citations;
  `);

  // Seed citations
  const citations = src.prepare('SELECT id, citation_hash, citation_text, created_at FROM citations').all();
  await seedTable('citations', citations, (row) => ({
    sql: `INSERT OR REPLACE INTO citations (id, citation_hash, citation_text, created_at) VALUES (?, ?, ?, ?)`,
    args: [row.id, row.citation_hash, row.citation_text, row.created_at]
  }));

  // Seed document_citations
  const docCitations = src.prepare('SELECT doc_id, citation_id, updated_at FROM document_citations').all();
  await seedTable('document_citations', docCitations, (row) => ({
    sql: `INSERT OR IGNORE INTO document_citations (doc_id, citation_id, updated_at) VALUES (?, ?, ?)`,
    args: [row.doc_id, row.citation_id, row.updated_at]
  }));

  // Seed catalogue_lookups
  let catalogueLookups = [];
  try {
    catalogueLookups = src.prepare('SELECT citation_id, hits, query_author, query_title, bib_id, looked_up_at FROM catalogue_lookups').all();
  } catch {
    console.log('  catalogue_lookups: table not found, skipping.');
  }
  await seedTable('catalogue_lookups', catalogueLookups, (row) => ({
    sql: `INSERT OR REPLACE INTO catalogue_lookups (citation_id, hits, query_author, query_title, bib_id, looked_up_at) VALUES (?, ?, ?, ?, ?, ?)`,
    args: [row.citation_id, row.hits ?? null, row.query_author ?? null, row.query_title ?? null, row.bib_id ?? null, row.looked_up_at]
  }));

  // Seed committee_members
  const committeeMembers = src.prepare('SELECT doc_id, name, role, affiliation, source, updated_at FROM committee_members').all();
  await seedTable('committee_members', committeeMembers, (row) => ({
    sql: `INSERT OR IGNORE INTO committee_members (doc_id, name, role, affiliation, source, updated_at) VALUES (?, ?, ?, ?, ?, ?)`,
    args: [row.doc_id, row.name, row.role ?? null, row.affiliation ?? null, row.source, row.updated_at]
  }));

  // Seed topics
  let topics = [];
  try {
    topics = src.prepare('SELECT topic_id, label, top_terms, doc_count, model_name, created_at FROM topics').all();
  } catch {
    console.log('  topics: table not found, skipping.');
  }
  await seedTable('topics', topics, (row) => ({
    sql: `INSERT OR REPLACE INTO topics (topic_id, label, top_terms, doc_count, model_name, created_at) VALUES (?, ?, ?, ?, ?, ?)`,
    args: [row.topic_id, row.label, row.top_terms, row.doc_count, row.model_name, row.created_at]
  }));

  // Seed document_topics
  let docTopics = [];
  try {
    docTopics = src.prepare('SELECT doc_id, topic_id, probability FROM document_topics').all();
  } catch {
    console.log('  document_topics: table not found, skipping.');
  }
  await seedTable('document_topics', docTopics, (row) => ({
    sql: `INSERT OR REPLACE INTO document_topics (doc_id, topic_id, probability) VALUES (?, ?, ?)`,
    args: [row.doc_id, row.topic_id, row.probability ?? null]
  }));

  // Seed document_topic_coords
  let topicCoords = [];
  try {
    topicCoords = src.prepare('SELECT doc_id, umap_x, umap_y FROM document_topic_coords').all();
  } catch {
    console.log('  document_topic_coords: table not found, skipping.');
  }
  await seedTable('document_topic_coords', topicCoords, (row) => ({
    sql: `INSERT OR REPLACE INTO document_topic_coords (doc_id, umap_x, umap_y) VALUES (?, ?, ?)`,
    args: [row.doc_id, row.umap_x, row.umap_y]
  }));

  // Seed topic_hierarchy_meta
  let hierarchyMeta = [];
  try {
    hierarchyMeta = src.prepare('SELECT id, leaf_topic_ids, linkage_json, created_at FROM topic_hierarchy_meta').all();
  } catch {
    console.log('  topic_hierarchy_meta: table not found, skipping.');
  }
  await seedTable('topic_hierarchy_meta', hierarchyMeta, (row) => ({
    sql: `INSERT OR REPLACE INTO topic_hierarchy_meta (id, leaf_topic_ids, linkage_json, created_at) VALUES (?, ?, ?, ?)`,
    args: [row.id, row.leaf_topic_ids, row.linkage_json, row.created_at]
  }));

  console.log('\nSeed complete.');
  src.close();
}

main().catch((err) => {
  console.error('Seed failed:', err);
  process.exit(1);
});
