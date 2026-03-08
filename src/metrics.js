import fs from 'node:fs';
import path from 'node:path';
import { DATA_DIR } from './config.js';
import {
  listAllDocumentMetadata,
  listAllFileMetrics,
  listAllCommitteeMembers,
  listAllCitationCounts,
  hasTopics,
  loadTopics,
  loadDocumentTopics,
  loadDocumentTopicCoords,
  loadTopicHierarchy,
  getCitationCooccurrence,
} from './db.js';
import {
  toArray, flattenText, extractYear, parsePageCount, topTermsFromText, buildWordCloud,
  buildMethodologyStats, extractNgrams, detectMethodologies, isLowSignalConceptPhrase
} from './nlp.js';
import { dedupeSupervisorNames } from './supervisors.js';
import { canonicalizeDomainText } from './domainDictionary.js';

function average(values) {
  if (!values.length) return null;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function stats(numbers) {
  if (!numbers.length) return { count: 0, min: null, max: null, mean: null };
  return {
    count: numbers.length,
    min: Math.min(...numbers),
    max: Math.max(...numbers),
    mean: average(numbers)
  };
}

function median(arr) {
  if (!arr.length) return null;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function firstPresent(obj, keys) {
  for (const key of keys) {
    const val = obj?.[key];
    if (val !== undefined && val !== null && val !== '') return val;
  }
  return null;
}

function unique(values) {
  return Array.from(new Set(values.filter(Boolean)));
}

function extractOcId(value) {
  const text = String(value || '').trim();
  if (!text) return null;

  const directId = text.match(/^\d+\.\d+$/);
  if (directId) return directId[0];

  const itemMatch = text.match(/\/items\/(\d+\.\d+)(?:[/?#]|$)/i);
  if (itemMatch) return itemMatch[1];

  const pdfMatch = text.match(/\/pdf\/\d+\/(\d+\.\d+)(?:[/?#]|$)/i);
  if (pdfMatch) return pdfMatch[1];

  return null;
}

function collectCandidateUrls(doc, rawId, doi) {
  const candidates = new Set();
  for (const val of [
    doc.uri, doc.URI, doc.isShownAt,
    doc.identifier, doc.Identifier,
    doi, rawId
  ]) {
    const text = String(val || '').trim();
    if (text) candidates.add(text);
  }
  for (const key of Object.keys(doc)) {
    if (key.toLowerCase().includes('url') || key.toLowerCase().includes('uri')) {
      const val = String(doc[key] || '').trim();
      if (val.startsWith('http')) candidates.add(val);
    }
  }
  return Array.from(candidates);
}

export function normalizeRecord(doc, dict = null) {
  const rawId = flattenText(firstPresent(doc, ['_id', 'id', 'identifier', 'Identifier'])) || '';
  const title = flattenText(firstPresent(doc, ['title', 'Title', 'name', 'Name']));
  const creators = toArray(firstPresent(doc, ['creator', 'Creator', 'author', 'Author']));
  const supervisors = dedupeSupervisorNames(toArray(firstPresent(doc, ['supervisor', 'Supervisor'])));
  const affiliation = toArray(firstPresent(doc, ['affiliation', 'Affiliation']));
  const dateRaw = firstPresent(doc, [
    'date_available', 'DateAvailable', 'dateAvailable',
    'dateIssued', 'DateIssued',
    'graduationDate', 'GraduationDate',
    'ubc_date_sort',
    'date', 'Date',
    'year', 'Year',
    'issued', 'Issued'
  ]);

  const description = flattenText(firstPresent(doc, ['description', 'Description', 'abstract', 'Abstract']));
  const fullText = flattenText(firstPresent(doc, ['full_text', 'FullText', 'transcript', 'text', 'ocr', 'body']));
  const subjects = toArray(firstPresent(doc, ['subject', 'Subject', 'subjects', 'keywords', 'keyword']));
  const program = toArray(firstPresent(doc, ['program_theses', 'program', 'Program']));
  const degree = toArray(firstPresent(doc, ['degree_theses', 'degree', 'Degree']));
  const genre = toArray(firstPresent(doc, ['genre', 'Genre']));
  const extentValues = toArray(firstPresent(doc, ['extent', 'Extent']));
  const uri = flattenText(firstPresent(doc, ['uri', 'URI', 'isShownAt', 'identifier', 'Identifier']));
  const rights = flattenText(firstPresent(doc, ['rights', 'Rights']));
  const doi = flattenText(firstPresent(doc, ['doi', 'DOI']));
  const campus = flattenText(firstPresent(doc, ['campus', 'Campus']));
  const scholarlyLevel = flattenText(firstPresent(doc, ['scholarly_level', 'scholarlyLevel', 'ScholarlyLevel']));
  const downloadCandidates = collectCandidateUrls(doc, rawId, doi);
  const derivedId = extractOcId(rawId) || extractOcId(uri) || downloadCandidates.map(extractOcId).find(Boolean);
  const stableId = derivedId || `${title}:${creators[0] || ''}`;

  const textForLength = fullText || description;
  const cleaned = String(textForLength).replace(/\s+/g, ' ').trim();
  const metadataWords = cleaned ? cleaned.split(' ').length : 0;
  const extentPages = parsePageCount(extentValues);
  const metadataPages = extentPages || Math.max(1, Math.round((metadataWords || 1) / 300));

  const themeText = [title, description, subjects.join(' '), program.join(' '), degree.join(' ')].join(' ');
  const themes = topTermsFromText(themeText, 12);
  const methodologies = detectMethodologies([title, description, subjects.join(' ')].join(' '));
  const conceptTerms = docConceptTerms({ title, abstract: description, subjects }, 12, dict);

  return {
    id: stableId,
    title,
    authors: creators,
    author: creators[0] || 'Unknown',
    supervisors,
    supervisorsSource: supervisors.length ? 'api' : null,
    affiliation,
    date: dateRaw ? String(dateRaw) : '',
    year: extractYear(dateRaw),
    degree: degree.join('; '),
    program: program.join('; '),
    type: genre.join('; '),
    rights,
    doi,
    campus,
    scholarlyLevel,
    extent: extentValues.join('; '),
    pages: metadataPages,
    pagesSource: extentPages ? 'metadata_extent' : 'estimated_from_metadata_words',
    abstract: description,
    subjects: subjects.length ? subjects : ['(Unspecified)'],
    wordCount: metadataWords,
    wordCountSource: 'metadata_text',
    charCount: cleaned.length,
    themes,
    methodologies,
    conceptTerms,
    uri,
    downloadCandidates,
    downloadUrl: null,
    downloadStatus: 'not_attempted',
    downloadError: null,
    fileBytes: null
  };
}

function buildMetrics(records, subjectLimit) {
  const conceptWords = new Map();
  const yearWords = new Map();
  const yearPages = new Map();

  for (const rec of records) {
    const concepts = Array.from(new Set((rec.conceptTerms || []).filter(Boolean)));
    if (concepts.length) {
      const weight = 1 / concepts.length;
      for (const concept of concepts) {
        if (!conceptWords.has(concept)) {
          conceptWords.set(concept, { weightedWordSum: 0, weightSum: 0, docCount: 0 });
        }
        const entry = conceptWords.get(concept);
        entry.weightedWordSum += rec.wordCount * weight;
        entry.weightSum += weight;
        entry.docCount += 1;
      }
    }

    if (rec.year) {
      if (!yearWords.has(rec.year)) yearWords.set(rec.year, []);
      if (!yearPages.has(rec.year)) yearPages.set(rec.year, []);
      yearWords.get(rec.year).push(rec.wordCount);
      yearPages.get(rec.year).push(rec.pages);
    }
  }

  const byConcept = Array.from(conceptWords.entries())
    .map(([concept, values]) => ({
      concept,
      docCount: values.docCount,
      weightedDocEquivalent: values.weightSum,
      weightedMean: values.weightSum ? (values.weightedWordSum / values.weightSum) : null
    }))
    .sort((a, b) => b.docCount - a.docCount || (b.weightedMean || 0) - (a.weightedMean || 0))
    .slice(0, subjectLimit);

  const byYear = Array.from(yearWords.entries())
    .map(([year, values]) => ({ year: Number(year), ...stats(values) }))
    .sort((a, b) => a.year - b.year);

  const avgPagesByYear = Array.from(yearPages.entries())
    .map(([year, values]) => ({ year: Number(year), ...stats(values) }))
    .sort((a, b) => a.year - b.year);

  const pageTrend = Array.from(yearPages.entries())
    .map(([year, values]) => ({
      year: Number(year),
      median: median(values),
      min: Math.min(...values),
      max: Math.max(...values),
      count: values.length
    }))
    .sort((a, b) => a.year - b.year);

  return {
    recordCount: records.length,
    overallWordCount: stats(records.map((r) => r.wordCount)),
    overallPageCount: stats(records.map((r) => r.pages)),
    overallCharCount: stats(records.map((r) => r.charCount)),
    byConcept,
    byYear,
    avgPagesByYear,
    pageTrend
  };
}

export function loadConceptDictionary() {
  try {
    const raw = fs.readFileSync(path.join(DATA_DIR, 'concepts', 'latest.json'), 'utf-8');
    const parsed = JSON.parse(raw);
    const canonicalSet = new Set((parsed.concepts || []).map((c) => c.canonical));
    const variantMap = parsed.variantToCanonical || {};
    return { canonicalSet, variantMap };
  } catch {
    return { canonicalSet: new Set(), variantMap: {} };
  }
}

export function docConceptTerms(rec, limit = 12, dict = null) {
  const { canonicalSet, variantMap } = dict || loadConceptDictionary();
  const text = [rec.title, rec.abstract, rec.subjects.join(' ')].join(' ');
  const seen = new Set();
  const result = [];
  for (const n of [2, 3]) {
    for (const ng of extractNgrams(text, n)) {
      const term = canonicalizeDomainText(ng);
      if (!term) continue;
      const canonical = variantMap[term] || (canonicalSet.has(term) ? term : null);
      if (!canonical || seen.has(canonical)) continue;
      seen.add(canonical);
      result.push(canonical);
      if (result.length >= limit) return result;
    }
  }
  return result;
}

function buildConceptCloud(records, maxTerms = 60) {
  const counts = new Map();
  for (const rec of records) {
    for (const term of (rec.conceptTerms || [])) {
      counts.set(term, (counts.get(term) || 0) + 1);
    }
  }
  return Array.from(counts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, maxTerms)
    .map(([term, count]) => ({ term, count }));
}

function buildSupervisorNgramMatrix(records, topN = 12, topM = 10) {
  const dict = loadConceptDictionary();
  const supervisorCounts = new Map();
  const docConceptCache = new Map();
  const supNgramCounts = new Map();

  for (const rec of records) {
    const terms = docConceptTerms(rec, 10, dict);
    const concepts = terms.map((label) => `c:${label.replace(/\s+/g, '_')}`);
    docConceptCache.set(rec.id, { concepts, labels: terms });
    const hasSup = rec.supervisors.length > 0;
    for (const sup of rec.supervisors) {
      supervisorCounts.set(sup, (supervisorCounts.get(sup) || 0) + 1);
    }
    if (hasSup) {
      for (const ng of concepts) {
        supNgramCounts.set(ng, (supNgramCounts.get(ng) || 0) + 1);
      }
    }
  }

  const topSupervisors = Array.from(supervisorCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, topN)
    .map(([name]) => name);
  const topNgrams = Array.from(supNgramCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, topM)
    .map(([conceptId]) => conceptId);

  const supSet = new Set(topSupervisors);
  const ngSet = new Set(topNgrams);

  const matrix = topSupervisors.map(() => topNgrams.map(() => 0));
  for (const rec of records) {
    const recSups = rec.supervisors.filter((s) => supSet.has(s));
    if (!recSups.length) continue;
    const recNgrams = (docConceptCache.get(rec.id)?.concepts || []).filter((ng) => ngSet.has(ng));
    for (const sup of recSups) {
      const si = topSupervisors.indexOf(sup);
      for (const ng of recNgrams) {
        const nj = topNgrams.indexOf(ng);
        matrix[si][nj] += 1;
      }
    }
  }

  const conceptIdToLabel = new Map();
  for (const { concepts, labels } of docConceptCache.values()) {
    for (let k = 0; k < concepts.length; k++) {
      conceptIdToLabel.set(concepts[k], labels[k]);
    }
  }

  return {
    supervisors: topSupervisors,
    ngrams: topNgrams.map((id) => conceptIdToLabel.get(id) || id),
    conceptIds: topNgrams,
    matrix
  };
}

function buildTermCooccurrence(records, topN = 20) {
  const dict = loadConceptDictionary();
  const pairCounts = new Map();
  for (const rec of records) {
    const concepts = docConceptTerms(rec, 15, dict);
    if (concepts.length < 2) continue;
    const sorted = [...concepts].sort();
    for (let i = 0; i < sorted.length; i++) {
      for (let j = i + 1; j < sorted.length; j++) {
        const key = `${sorted[i]}|||${sorted[j]}`;
        pairCounts.set(key, (pairCounts.get(key) || 0) + 1);
      }
    }
  }
  return Array.from(pairCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, topN)
    .map(([key, count]) => {
      const [termA, termB] = key.split('|||');
      return {
        conceptIdA: `c:${termA.replace(/\s+/g, '_')}`,
        conceptIdB: `c:${termB.replace(/\s+/g, '_')}`,
        termA,
        termB,
        count
      };
    });
}

function buildConceptTimeline(records, topN = 8) {
  const dict = loadConceptDictionary();
  const conceptDocCounts = new Map();
  const conceptYearCounts = new Map();

  for (const rec of records) {
    const concepts = docConceptTerms(rec, 12, dict);
    for (const concept of concepts) {
      conceptDocCounts.set(concept, (conceptDocCounts.get(concept) || 0) + 1);
      if (rec.year) {
        if (!conceptYearCounts.has(concept)) conceptYearCounts.set(concept, new Map());
        const yearMap = conceptYearCounts.get(concept);
        yearMap.set(rec.year, (yearMap.get(rec.year) || 0) + 1);
      }
    }
  }

  const topConcepts = Array.from(conceptDocCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, topN)
    .map(([concept]) => concept);

  return topConcepts.map((concept) => {
    const yearMap = conceptYearCounts.get(concept) || new Map();
    const data = Array.from(yearMap.entries())
      .map(([year, count]) => ({ year: Number(year), count }))
      .sort((a, b) => a.year - b.year);
    return {
      concept,
      totalDocs: conceptDocCounts.get(concept) || 0,
      data
    };
  });
}

function buildMethodologyConceptMatrix(records, topM = 10, topC = 10) {
  const dict = loadConceptDictionary();
  const methodCounts = new Map();
  const conceptCounts = new Map();
  const docConceptCache = new Map();

  for (const rec of records) {
    const terms = docConceptTerms(rec, 10, dict);
    const concepts = terms.map((label) => `c:${label.replace(/\s+/g, '_')}`);
    docConceptCache.set(rec.id, { concepts, labels: terms });

    for (const m of (rec.methodologies || [])) {
      methodCounts.set(m, (methodCounts.get(m) || 0) + 1);
      for (const c of concepts) {
        conceptCounts.set(c, (conceptCounts.get(c) || 0) + 1);
      }
    }
  }

  const topMethodologies = Array.from(methodCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, topM)
    .map(([name]) => name);
  const topConcepts = Array.from(conceptCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, topC)
    .map(([conceptId]) => conceptId);

  const methSet = new Set(topMethodologies);
  const conSet = new Set(topConcepts);

  const matrix = topMethodologies.map(() => topConcepts.map(() => 0));
  for (const rec of records) {
    const recMeths = (rec.methodologies || []).filter((m) => methSet.has(m));
    if (!recMeths.length) continue;
    const recConcepts = (docConceptCache.get(rec.id)?.concepts || []).filter((c) => conSet.has(c));
    for (const m of recMeths) {
      const mi = topMethodologies.indexOf(m);
      for (const c of recConcepts) {
        const ci = topConcepts.indexOf(c);
        matrix[mi][ci] += 1;
      }
    }
  }

  const conceptIdToLabel = new Map();
  for (const { concepts, labels } of docConceptCache.values()) {
    for (let k = 0; k < concepts.length; k++) {
      conceptIdToLabel.set(concepts[k], labels[k]);
    }
  }

  return {
    methodologies: topMethodologies,
    concepts: topConcepts.map((id) => conceptIdToLabel.get(id) || id),
    conceptIds: topConcepts,
    matrix
  };
}

function buildResearchGaps(records, topN = 15) {
  const dict = loadConceptDictionary();
  const conceptDocCounts = new Map();
  const cooccurrenceCounts = new Map();

  for (const rec of records) {
    const concepts = docConceptTerms(rec, 15, dict);
    for (const c of concepts) {
      conceptDocCounts.set(c, (conceptDocCounts.get(c) || 0) + 1);
    }
    if (concepts.length >= 2) {
      const sorted = [...concepts].sort();
      for (let i = 0; i < sorted.length; i++) {
        for (let j = i + 1; j < sorted.length; j++) {
          const key = `${sorted[i]}|||${sorted[j]}`;
          cooccurrenceCounts.set(key, (cooccurrenceCounts.get(key) || 0) + 1);
        }
      }
    }
  }

  const topConcepts = Array.from(conceptDocCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 20)
    .map(([c]) => c);

  const gaps = [];
  for (let i = 0; i < topConcepts.length; i++) {
    for (let j = i + 1; j < topConcepts.length; j++) {
      const a = topConcepts[i];
      const b = topConcepts[j];
      const key = a < b ? `${a}|||${b}` : `${b}|||${a}`;
      const cooccurrence = cooccurrenceCounts.get(key) || 0;
      const countA = conceptDocCounts.get(a) || 0;
      const countB = conceptDocCounts.get(b) || 0;
      const gapScore = (countA * countB) / (cooccurrence + 1);
      gaps.push({ conceptA: a, conceptB: b, countA, countB, cooccurrence, gapScore });
    }
  }

  return gaps
    .sort((a, b) => b.gapScore - a.gapScore)
    .slice(0, topN);
}

function buildParentClusters(hierarchy, topics, targetK = 10) {
  const { leafTopicIds, linkage } = hierarchy;
  const N = leafTopicIds.length;
  if (N <= targetK) {
    const parentClusters = topics.filter((t) => t.topicId !== -1).map((t, i) => ({
      parentId: i, topicId: i, label: t.label, docCount: t.docCount, children: [t.topicId],
    }));
    const leafToParent = new Map(parentClusters.map((p) => [p.children[0], p.parentId]));
    return { parentClusters, leafToParent };
  }

  const parent = new Array(N + linkage.length).fill(-1);
  const find = (x) => { while (parent[x] !== -1) x = parent[x]; return x; };
  const union = (a, b, into) => { parent[find(a)] = into; parent[find(b)] = into; };

  const mergesToApply = N - targetK;
  for (let m = 0; m < mergesToApply; m++) {
    const [i, j] = linkage[m];
    union(i, j, N + m);
  }

  const rootToChildren = new Map();
  const topicDocCount = new Map(topics.map((t) => [t.topicId, t.docCount || 0]));
  for (let i = 0; i < N; i++) {
    const root = find(i);
    if (!rootToChildren.has(root)) rootToChildren.set(root, []);
    rootToChildren.get(root).push(leafTopicIds[i]);
  }

  const topicLabelMap = new Map(topics.map((t) => [t.topicId, t.label]));
  const parentClusters = [];
  const leafToParent = new Map();
  let parentIdx = 0;

  for (const [, children] of rootToChildren) {
    const bestChild = children.reduce((best, tid) =>
      (topicDocCount.get(tid) || 0) > (topicDocCount.get(best) || 0) ? tid : best
    , children[0]);
    const totalDocs = children.reduce((sum, tid) => sum + (topicDocCount.get(tid) || 0), 0);
    parentClusters.push({
      parentId: parentIdx, topicId: parentIdx,
      label: topicLabelMap.get(bestChild) || `Cluster ${parentIdx}`,
      docCount: totalDocs, children,
    });
    for (const tid of children) leafToParent.set(tid, parentIdx);
    parentIdx++;
  }

  parentClusters.sort((a, b) => b.docCount - a.docCount);
  const oldToNew = new Map(parentClusters.map((p, i) => [p.parentId, i]));
  for (const p of parentClusters) p.parentId = p.topicId = oldToNew.get(p.parentId);
  for (const [tid, oldPid] of leafToParent) leafToParent.set(tid, oldToNew.get(oldPid));

  return { parentClusters, leafToParent };
}

function buildTopicsByYear(topics, documents, leafToParent) {
  const yearCounts = new Map();
  for (const doc of documents) {
    if (doc.topicId == null || !doc.year) continue;
    const resolvedId = leafToParent ? (leafToParent.get(doc.topicId) ?? doc.topicId) : doc.topicId;
    if (!yearCounts.has(resolvedId)) yearCounts.set(resolvedId, new Map());
    const ym = yearCounts.get(resolvedId);
    ym.set(doc.year, (ym.get(doc.year) || 0) + 1);
  }
  return topics
    .filter((t) => t.topicId !== -1)
    .slice(0, 10)
    .map((topic) => {
      const ym = yearCounts.get(topic.topicId) || new Map();
      const data = Array.from(ym.entries())
        .map(([year, count]) => ({ year: Number(year), count }))
        .sort((a, b) => a.year - b.year);
      return { topicId: topic.topicId, label: topic.label, data };
    });
}

function buildSupervisorNetwork(records, minEdge = 1) {
  const nodeMap = new Map();
  const edgeMap = new Map();

  for (const rec of records) {
    const people = [];
    for (const s of (rec.supervisors || [])) people.push(s);
    if (rec.committee) {
      for (const m of rec.committee) {
        if (m.role === 'Supervisor' || m.role === 'Co-Supervisor' || m.role === 'Committee Member') {
          if (!people.includes(m.name)) people.push(m.name);
        }
      }
    }
    if (people.length === 0) continue;
    for (const p of people) {
      nodeMap.set(p, (nodeMap.get(p) || 0) + 1);
    }
    const sorted = [...people].sort();
    for (let i = 0; i < sorted.length; i++) {
      for (let j = i + 1; j < sorted.length; j++) {
        const key = `${sorted[i]}|||${sorted[j]}`;
        if (!edgeMap.has(key)) edgeMap.set(key, { weight: 0, docs: [] });
        const e = edgeMap.get(key);
        e.weight += 1;
        e.docs.push(rec.id);
      }
    }
  }

  const topNodes = Array.from(nodeMap.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30);
  const topSet = new Set(topNodes.map(([name]) => name));

  const nodes = topNodes.map(([id, docCount]) => ({ id, docCount }));
  const edges = Array.from(edgeMap.entries())
    .filter(([key, e]) => {
      const [a, b] = key.split('|||');
      return topSet.has(a) && topSet.has(b) && e.weight >= minEdge;
    })
    .map(([key, e]) => {
      const [source, target] = key.split('|||');
      return { source, target, weight: e.weight, docs: e.docs };
    });

  return { nodes, edges };
}

function buildCitationCooccurrenceNetwork(rows) {
  const nodeMap = new Map();
  for (const row of rows) {
    if (!nodeMap.has(row.id1)) nodeMap.set(row.id1, { id: row.id1, label: row.text1, freq: Number(row.freq1) });
    if (!nodeMap.has(row.id2)) nodeMap.set(row.id2, { id: row.id2, label: row.text2, freq: Number(row.freq2) });
  }
  const nodes = Array.from(nodeMap.values());
  const edges = rows.map(row => ({
    source: row.id1,
    target: row.id2,
    weight: Number(row.shared)
  }));
  return { nodes, edges };
}

function buildMethodologyTopicMatrix(records, topics) {
  const methCounts = new Map();
  for (const rec of records) {
    for (const m of (rec.methodologies || [])) {
      methCounts.set(m, (methCounts.get(m) || 0) + 1);
    }
  }
  const topMethodologies = Array.from(methCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([name]) => name);

  const validTopics = topics.filter(t => t.topicId !== -1).slice(0, 8);
  const methSet = new Set(topMethodologies);
  const topicIdList = validTopics.map(t => t.topicId);
  const topicIdSet = new Set(topicIdList);
  const matrix = topMethodologies.map(() => topicIdList.map(() => 0));

  for (const rec of records) {
    if (rec.topicId == null || !topicIdSet.has(rec.topicId)) continue;
    const recMeths = (rec.methodologies || []).filter(m => methSet.has(m));
    const ti = topicIdList.indexOf(rec.topicId);
    for (const m of recMeths) {
      const mi = topMethodologies.indexOf(m);
      matrix[mi][ti] += 1;
    }
  }

  return {
    methodologies: topMethodologies,
    topics: validTopics.map(t => ({ topicId: t.topicId, label: t.label })),
    matrix
  };
}

export async function collectMetricsFromDb({ subjectLimit = 25 } = {}) {
  const conceptDict = loadConceptDictionary();
  const [allDocs, fileMetricsMap, committeeMembersMap, citationCountsMap] = await Promise.all([
    listAllDocumentMetadata(),
    listAllFileMetrics(),
    listAllCommitteeMembers(),
    listAllCitationCounts(),
  ]);

  const records = allDocs.map(({ docId, metadata }) => {
    const rec = normalizeRecord(metadata, conceptDict);
    const fm = fileMetricsMap.get(docId);
    if (fm) {
      if (fm.word_count) { rec.wordCount = Number(fm.word_count); rec.wordCountSource = fm.word_source || 'pdf'; }
      if (fm.page_count) { rec.pages = Number(fm.page_count); rec.pagesSource = fm.page_source || 'pdf'; }
      rec.downloadUrl = fm.download_url || null;
      rec.downloadStatus = fm.status || 'not_attempted';
      rec.fileBytes = fm.file_bytes ? Number(fm.file_bytes) : null;
    }
    const committee = committeeMembersMap.get(docId) || [];
    if (committee.length) {
      const supervisorNames = committee.map((m) => m.name).filter(Boolean);
      if (supervisorNames.length) rec.supervisors = dedupeSupervisorNames(supervisorNames);
      rec.committee = committee.map((m) => ({
        name: m.name,
        role: m.role || 'Committee Member',
        affiliation: m.affiliation || null,
        source: m.source || 'unknown',
      }));
    }
    rec.citationCount = citationCountsMap.get(docId) || 0;
    return rec;
  });

  // Load BERTopic results if available
  let topicData = null;
  const topicsExist = await hasTopics();
  if (topicsExist) {
    const [topics, docTopics, topicCoords, hierarchy] = await Promise.all([
      loadTopics(),
      loadDocumentTopics(),
      loadDocumentTopicCoords(),
      loadTopicHierarchy(),
    ]);
    for (const doc of records) {
      const dt = docTopics.get(doc.id);
      if (dt) {
        doc.topicId = dt.topicId;
        doc.topicProbability = dt.probability;
      }
      const coord = topicCoords.get(doc.id);
      if (coord) {
        doc.umapX = coord.x;
        doc.umapY = coord.y;
      }
    }
    const parentInfo = hierarchy ? buildParentClusters(hierarchy, topics) : null;
    topicData = {
      topics,
      byYear: parentInfo
        ? buildTopicsByYear(parentInfo.parentClusters, records, parentInfo.leafToParent)
        : buildTopicsByYear(topics, records),
      hierarchy,
    };
  }

  // Citation cooccurrence network
  const citationCoocRows = await getCitationCooccurrence(100);

  const metrics = buildMetrics(records, subjectLimit);
  return {
    generatedAt: new Date().toISOString(),
    metrics,
    documents: records,
    wordCloud: buildWordCloud(records),
    ngramCloud: buildConceptCloud(records),
    methodologies: buildMethodologyStats(records),
    supervisorNgramMatrix: buildSupervisorNgramMatrix(records),
    termCooccurrence: buildTermCooccurrence(records),
    conceptTimeline: buildConceptTimeline(records),
    methodologyConceptMatrix: buildMethodologyConceptMatrix(records),
    researchGaps: buildResearchGaps(records),
    supervisorNetwork: buildSupervisorNetwork(records),
    citationCooccurrence: buildCitationCooccurrenceNetwork(citationCoocRows),
    methodologyTopicMatrix: topicData ? buildMethodologyTopicMatrix(records, topicData.topics) : null,
    topicData,
  };
}
