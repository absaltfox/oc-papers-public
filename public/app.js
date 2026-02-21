// --- DOM references ---
const statusTextEl = document.getElementById('statusText');
const spinnerEl = document.getElementById('spinner');
const statusEl = document.getElementById('status');
const documentsTableEl = document.getElementById('documentsTable');
const docFilterEl = document.getElementById('docFilter');
const docTheadRow = document.querySelector('#tab-records thead tr');
const selectAllDocsEl = document.getElementById('selectAllDocs');
const docDetailsEl = document.getElementById('docDetails');
const kpisEl = document.getElementById('kpis');
const pagesByYearChartEl = document.getElementById('pagesByYearChart');
const wordCloudEl = document.getElementById('wordCloud');
const themeResultsEl = document.getElementById('themeResults');
const subjectBarsEl = document.getElementById('subjectBars');
const dissertationsByYearChartEl = document.getElementById('dissertationsByYearChart');
const wordsByYearChartEl = document.getElementById('wordsByYearChart');
const pageTrendChartEl = document.getElementById('pageTrendChart');
const ngramCloudEl = document.getElementById('ngramCloud');
const methodologyBarsEl = document.getElementById('methodologyBars');
const cooccurrenceBarsEl = document.getElementById('cooccurrenceBars');
const supervisorHeatmapEl = document.getElementById('supervisorHeatmap');
const conceptTimelineChartEl = document.getElementById('conceptTimelineChart');
const conceptTimelineLegendEl = document.getElementById('conceptTimelineLegend');
const methodologyConceptHeatmapEl = document.getElementById('methodologyConceptHeatmap');
const researchGapsListEl = document.getElementById('researchGapsList');
const foundationalWorksListEl = document.getElementById('foundationalWorksList');
const exportBibTeXBtn = document.getElementById('exportBibTeX');
const exportRISBtn = document.getElementById('exportRIS');
const exportCitationBibTeXBtn = document.getElementById('exportCitationBibTeX');
const exportCitationRISBtn = document.getElementById('exportCitationRIS');
const tabButtons = Array.from(document.querySelectorAll('.tab-btn'));
const tabPanels = Array.from(document.querySelectorAll('.tab-panel'));

// Modal elements
const docModalOverlay = document.getElementById('docModalOverlay');
const docModalCloseBtn = document.getElementById('docModalClose');
const docModalTitleEl = document.getElementById('docModalTitle');

// Citation Explorer elements
const citationDocsTableEl = document.getElementById('citationDocsTable');
const citationDocFilterEl = document.getElementById('citationDocFilter');
const citationListTitleEl = document.getElementById('citationListTitle');
const citationEntriesEl = document.getElementById('citationEntries');
const citationTabButtons = Array.from(document.querySelectorAll('.citation-tab-btn'));

// --- State ---
const state = {
  payload: null,
  selectedDocId: null,
  selectedTheme: null,
  loading: false,
  sortKey: null,
  sortDir: 'asc',
  filterText: '',
  citationDocId: null,
  citationFilterText: '',
  citationRequestToken: 0,
  selectedDocIds: new Set(),
  selectedCitationIds: new Set(),
};

// --- Utilities ---

function formatNum(value) {
  if (value === null || value === undefined) return '-';
  return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(value);
}

function formatBytes(bytes) {
  if (!bytes) return '-';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function escapeHtml(text) {
  return String(text || '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function setStatus(message, isError = false) {
  statusTextEl.textContent = message;
  statusEl.classList.toggle('error', isError);
}

function showSpinner(show) {
  spinnerEl.hidden = !show;
}

// --- Tab navigation ---

function setActiveTab(tabName) {
  for (const btn of tabButtons) {
    const isActive = btn.dataset.tab === tabName;
    btn.classList.toggle('active', isActive);
    btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
  }
  for (const panel of tabPanels) {
    panel.classList.toggle('active', panel.id === `tab-${tabName}`);
  }
  if (tabName === 'citations' && state.payload) {
    renderCitationDocs();
    setActiveCitationTab('browse');
  }
}

function setActiveCitationTab(tabName) {
  for (const btn of citationTabButtons) {
    btn.classList.toggle('active', btn.dataset.citationTab === tabName);
  }
  for (const section of document.querySelectorAll('.citation-tab-section')) {
    section.classList.toggle('active', section.id === `citation-${tabName}`);
  }
  if (tabName === 'foundational' && state.payload) {
    loadFoundationalWorks();
  }
}

// --- Document rendering ---

function intersectionCount(a, b) {
  const setB = new Set(b);
  let count = 0;
  for (const x of a) {
    if (setB.has(x)) count += 1;
  }
  return count;
}

function relatedDocuments(doc, allDocs, limit = 6) {
  const docThemes = doc.themes || [];
  return allDocs
    .filter((candidate) => candidate.id !== doc.id)
    .map((candidate) => {
      const overlap = intersectionCount(docThemes, candidate.themes || []);
      return {
        ...candidate,
        overlap,
        sharedThemes: (candidate.themes || []).filter((t) => docThemes.includes(t)).slice(0, 4)
      };
    })
    .filter((item) => item.overlap > 0)
    .sort((a, b) => b.overlap - a.overlap || (b.year || 0) - (a.year || 0))
    .slice(0, limit);
}

function openRecord(docId, focusTab = 'records') {
  state.selectedDocId = docId;
  renderDocuments();
  renderDetails();
  docModalOverlay.hidden = false;
  setActiveTab(focusTab);
}

function closeDocModal() {
  docModalOverlay.hidden = true;
}

function docsForTheme(theme) {
  const docs = state.payload?.documents || [];
  const normalized = String(theme || '').toLowerCase();
  return docs.filter((doc) => (doc.themes || []).some((t) => t.toLowerCase() === normalized));
}

function docsForConceptTerm(term) {
  const docs = state.payload?.documents || [];
  const normalized = String(term || '').toLowerCase();
  return docs.filter((doc) => (doc.conceptTerms || []).some((t) => String(t || '').toLowerCase() === normalized));
}

function docsForMethodology(methodology) {
  const docs = state.payload?.documents || [];
  const normalized = String(methodology || '').toLowerCase();
  return docs.filter((doc) => (doc.methodologies || []).some((m) => String(m || '').toLowerCase() === normalized));
}

function docsForCooccurrence(termA, termB) {
  const a = String(termA || '').toLowerCase();
  const b = String(termB || '').toLowerCase();
  if (!a || !b) return [];
  const docs = state.payload?.documents || [];
  return docs.filter((doc) => {
    const terms = new Set((doc.conceptTerms || []).map((t) => String(t || '').toLowerCase()));
    return terms.has(a) && terms.has(b);
  });
}

function docsForSupervisorConcept(supervisor, concept) {
  const sup = String(supervisor || '').toLowerCase();
  const conceptNorm = String(concept || '').toLowerCase();
  const docs = state.payload?.documents || [];
  return docs.filter((doc) => {
    const hasSup = (doc.supervisors || []).some((s) => String(s || '').toLowerCase() === sup);
    if (!hasSup) return false;
    const terms = new Set((doc.conceptTerms || []).map((t) => String(t || '').toLowerCase()));
    return terms.has(conceptNorm);
  });
}

function openMatchesModal(title, matches) {
  const list = matches || [];
  const body = list.length
    ? `
      <div class="related-list">
        ${list
          .map(
            (doc) => `
            <div class="related-item" data-related-id="${escapeHtml(doc.id)}">
              <strong>${escapeHtml(doc.title || '(Untitled)')}</strong>
              <p>${escapeHtml(doc.author || 'Unknown')} &middot; ${doc.year || '-'} &middot; ${escapeHtml(doc.degree || '-')}</p>
            </div>
          `
          )
          .join('')}
      </div>
    `
    : '<p class="meta">No matching dissertations found in the current result set.</p>';

  docDetailsEl.innerHTML = `
    <div class="meta">
      <p><strong>${escapeHtml(title)}</strong></p>
      <p>${formatNum(list.length)} dissertation(s)</p>
    </div>
    ${body}
  `;

  for (const item of docDetailsEl.querySelectorAll('.related-item[data-related-id]')) {
    item.addEventListener('click', () => {
      const targetId = item.getAttribute('data-related-id');
      if (targetId) openRecord(targetId, 'records');
    });
  }
  docModalOverlay.hidden = false;
}

function docSortValue(doc, key) {
  switch (key) {
    case 'title': return (doc.title || '').toLowerCase();
    case 'author': return (doc.author || '').toLowerCase();
    case 'year': return doc.year || 0;
    case 'degree': return (doc.degree || doc.type || '').toLowerCase();
    case 'pages': return doc.pages || 0;
    case 'wordCount': return doc.wordCount || 0;
    default: return '';
  }
}

function getFilteredSortedDocs() {
  let docs = state.payload?.documents || [];

  if (state.filterText) {
    const q = state.filterText.toLowerCase();
    docs = docs.filter((doc) =>
      (doc.title || '').toLowerCase().includes(q) ||
      (doc.author || '').toLowerCase().includes(q) ||
      (doc.supervisors || []).some((name) => String(name || '').toLowerCase().includes(q)) ||
      (doc.degree || '').toLowerCase().includes(q) ||
      (doc.program || '').toLowerCase().includes(q) ||
      String(doc.year || '').includes(q)
    );
  }

  if (state.sortKey) {
    const dir = state.sortDir === 'asc' ? 1 : -1;
    docs = [...docs].sort((a, b) => {
      const av = docSortValue(a, state.sortKey);
      const bv = docSortValue(b, state.sortKey);
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      if (av < bv) return -1 * dir;
      if (av > bv) return 1 * dir;
      return 0;
    });
  }

  return docs;
}

function updateSortHeaders() {
  for (const th of docTheadRow.querySelectorAll('th.sortable')) {
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.dataset.sortKey === state.sortKey) {
      th.classList.add(state.sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
    }
  }
}

function renderDocuments() {
  const docs = getFilteredSortedDocs();

  documentsTableEl.innerHTML = docs
    .map((doc) => {
      const active = doc.id === state.selectedDocId ? ' active' : '';
      const checked = state.selectedDocIds.has(doc.id) ? ' checked' : '';
      return `
        <tr class="doc-row${active}" data-doc-id="${escapeHtml(doc.id)}">
          <td class="doc-check-col"><input type="checkbox" class="doc-row-check" data-check-id="${escapeHtml(doc.id)}"${checked} /></td>
          <td>${escapeHtml(doc.title || '(Untitled)')}</td>
          <td>${escapeHtml(doc.author || '')}</td>
          <td>${doc.year || '-'}</td>
          <td>${escapeHtml(doc.degree || doc.type || '-')}</td>
          <td>${formatNum(doc.pages)}</td>
          <td>${formatNum(doc.wordCount)}</td>
        </tr>
      `;
    })
    .join('');

  for (const row of documentsTableEl.querySelectorAll('.doc-row')) {
    row.addEventListener('click', (e) => {
      if (e.target.closest('.doc-row-check')) return;
      openRecord(row.dataset.docId, 'records');
    });
  }

  for (const cb of documentsTableEl.querySelectorAll('.doc-row-check')) {
    cb.addEventListener('change', () => {
      const id = cb.dataset.checkId;
      if (cb.checked) state.selectedDocIds.add(id);
      else state.selectedDocIds.delete(id);
      syncSelectAllDocs();
    });
  }

  syncSelectAllDocs();
  updateSortHeaders();
}

function syncSelectAllDocs() {
  const visibleChecks = documentsTableEl.querySelectorAll('.doc-row-check');
  const allChecked = visibleChecks.length > 0 && Array.from(visibleChecks).every((cb) => cb.checked);
  selectAllDocsEl.checked = allChecked;
  selectAllDocsEl.indeterminate = !allChecked && state.selectedDocIds.size > 0;
}

function renderDetails() {
  const docs = state.payload?.documents || [];
  if (!docs.length) {
    docModalTitleEl.textContent = 'Document Details';
    docDetailsEl.textContent = 'No documents in current result set.';
    return;
  }

  let doc = docs.find((d) => d.id === state.selectedDocId);
  if (!doc) {
    doc = docs[0];
    state.selectedDocId = doc.id;
  }

  docModalTitleEl.textContent = doc.title || '(Untitled)';

  const related = relatedDocuments(doc, docs);
  const abstract = doc.abstract
    ? doc.abstract.split(/\n{2,}|\r?\n/).map((p) => `<p>${escapeHtml(p.trim())}</p>`).join('')
    : '<p>No abstract provided.</p>';
  const themes = doc.themes?.length
    ? doc.themes.map((t) => `<span class="token">${escapeHtml(t)}</span>`).join('')
    : '<span class="token">No themes</span>';
  const concepts = doc.conceptTerms?.length
    ? doc.conceptTerms.map((t) => `<span class="token concept">${escapeHtml(t)}</span>`).join('')
    : '<span class="token concept">No concepts</span>';

  const relatedHtml = related.length
    ? related
        .map(
          (r) => `
          <div class="related-item" data-related-id="${escapeHtml(r.id)}">
            <strong>${escapeHtml(r.title || '(Untitled)')}</strong>
            <p>${escapeHtml(r.author || 'Unknown')} &middot; ${r.year || '-'} &middot; Shared themes: ${escapeHtml(r.sharedThemes.join(', '))}</p>
          </div>
        `
        )
        .join('')
    : '<p class="meta">No related documents identified from overlapping themes.</p>';

  let committeeHtml = '';
  const hasSupervisorsShown = doc.supervisors?.length > 0;
  const supervisorRoles = new Set(['Supervisor', 'Co-Supervisor']);
  if (doc.committee?.length) {
    const grouped = {};
    for (const m of doc.committee) {
      const role = m.role || 'Committee Member';
      if (hasSupervisorsShown && supervisorRoles.has(role)) continue;
      if (!grouped[role]) grouped[role] = [];
      grouped[role].push(m);
    }
    committeeHtml = Object.entries(grouped).map(([role, members]) =>
      members.map((m) =>
        `<div class="detail-meta-label">${escapeHtml(role)}</div><div class="detail-meta-value">${escapeHtml(m.name)}${m.affiliation ? ` (${escapeHtml(m.affiliation)})` : ''}</div>`
      ).join('')
    ).join('');
  }

  const subtitleParts = [
    doc.author || 'Unknown',
    doc.year || '',
    doc.degree || ''
  ].filter(Boolean);

  const actions = [];
  if (doc.downloadUrl) {
    actions.push(`<a class="btn ghost btn-sm" href="${escapeHtml(doc.downloadUrl)}" target="_blank" rel="noreferrer">Open PDF</a>`);
  }
  if (doc.uri) {
    actions.push(`<a class="btn ghost btn-sm" href="${escapeHtml(doc.uri)}" target="_blank" rel="noreferrer">Open Record</a>`);
  }
  actions.push(`<button class="btn ghost btn-sm" data-doc-bibtex>BibTeX</button>`);
  const actionsHtml = `<div class="doc-actions">${actions.join('')}</div>`;
  const downloadNoteHtml = doc.downloadError
    ? `<p class="detail-download-note">${escapeHtml(doc.downloadError)}</p>`
    : '';

  const citationCount = doc.citationCount || 0;
  const citationsHtml = citationCount > 0
    ? `<details class="citations-details" data-doc-id="${escapeHtml(doc.id)}">
        <summary>Works Cited (${formatNum(citationCount)} references)</summary>
        <div class="citations-content"><p class="meta">Loading...</p></div>
      </details>`
    : '';

  docDetailsEl.innerHTML = `
    <p class="doc-subtitle">${escapeHtml(subtitleParts.join(' \u00B7 '))}</p>
    ${actionsHtml}
    ${downloadNoteHtml}
    <div class="detail-meta">
      <div class="detail-meta-label">Date</div><div class="detail-meta-value">${escapeHtml((doc.date || '-').replace(/\s*AD\s*$/i, ''))}</div>
      <div class="detail-meta-label">Program</div><div class="detail-meta-value">${escapeHtml(doc.program || '-')}</div>
      <div class="detail-meta-label">Pages</div><div class="detail-meta-value">${formatNum(doc.pages)}</div>
      <div class="detail-meta-label">Words</div><div class="detail-meta-value">${formatNum(doc.wordCount)}</div>
      ${doc.supervisors?.length ? `<div class="detail-meta-label">Supervisor</div><div class="detail-meta-value">${doc.supervisors.map((s) => `<button class="supervisor-link" data-supervisor-name="${escapeHtml(s)}">${escapeHtml(s)}</button>`).join('; ')}</div>` : ''}
      ${committeeHtml}
    </div>
    <div>
      <p class="detail-section-title">Abstract</p>
      <div class="detail-abstract">${abstract}</div>
    </div>
    <div>
      <p class="detail-section-title">Key Themes</p>
      <div class="token-list">${themes}</div>
    </div>
    <div>
      <p class="detail-section-title">Concepts</p>
      <div class="token-list">${concepts}</div>
    </div>
    <div>
      <p class="detail-section-title">Related Documents</p>
      <div class="related-list">${relatedHtml}</div>
    </div>
    ${citationsHtml}
  `;

  for (const item of docDetailsEl.querySelectorAll('.related-item[data-related-id]')) {
    item.addEventListener('click', () => {
      const targetId = item.getAttribute('data-related-id');
      if (targetId) openRecord(targetId, 'records');
    });
  }

  for (const btn of docDetailsEl.querySelectorAll('.supervisor-link[data-supervisor-name]')) {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openSupervisorProfile(btn.dataset.supervisorName);
    });
  }

  const bibBtn = docDetailsEl.querySelector('[data-doc-bibtex]');
  if (bibBtn) {
    bibBtn.addEventListener('click', () => {
      downloadFile(generateBibTeX([doc]), `${sanitizeBibKey(doc.author)}${doc.year || ''}.bib`, 'application/x-bibtex');
    });
  }

  const citationsDetails = docDetailsEl.querySelector('.citations-details[data-doc-id]');
  if (citationsDetails) {
    citationsDetails.addEventListener('toggle', async () => {
      if (!citationsDetails.open) return;
      const contentEl = citationsDetails.querySelector('.citations-content');
      if (contentEl.dataset.loaded) return;
      contentEl.dataset.loaded = '1';
      try {
        const docId = citationsDetails.dataset.docId;
        const res = await fetch(`/api/documents/${encodeURIComponent(docId)}/citations`);
        if (!res.ok) {
          contentEl.innerHTML = '<p class="meta">Failed to load citations.</p>';
          return;
        }
        const data = await res.json();
        if (!data.citations?.length) {
          contentEl.innerHTML = '<p class="meta">No citations found.</p>';
          return;
        }
        contentEl.innerHTML = data.citations.map((c) =>
          `<p class="citation-entry">${escapeHtml(c.citation_text)}${catalogueBadge(c)}</p>`
        ).join('');
      } catch {
        contentEl.innerHTML = '<p class="meta">Connection error.</p>';
      }
    });
  }
}

// --- Analytics rendering ---

function renderKpis() {
  const metrics = state.payload?.metrics;
  if (!metrics) {
    kpisEl.innerHTML = '';
    return;
  }

  const docs = state.payload?.documents || [];
  const citeCounts = docs.map((d) => d.citationCount || 0).filter((c) => c > 0);
  const citeMin = citeCounts.length ? Math.min(...citeCounts) : null;
  const citeMax = citeCounts.length ? Math.max(...citeCounts) : null;
  const citeMean = citeCounts.length
    ? citeCounts.reduce((a, b) => a + b, 0) / citeCounts.length
    : null;

  const cards = [
    { label: 'Retrieved Records', value: metrics.recordCount },
    { label: 'Mean Pages', value: metrics.overallPageCount.mean },
    { label: 'Min Pages', value: metrics.overallPageCount.min },
    { label: 'Max Pages', value: metrics.overallPageCount.max },
    { label: 'Mean Words', value: metrics.overallWordCount.mean },
    { label: 'Min Words', value: metrics.overallWordCount.min },
    { label: 'Max Words', value: metrics.overallWordCount.max },
    { label: 'Mean Works Cited', value: citeMean },
    { label: 'Min Works Cited', value: citeMin },
    { label: 'Max Works Cited', value: citeMax }
  ];

  kpisEl.innerHTML = cards
    .map(
      (card) => `
      <article class="kpi">
        <p>${card.label}</p>
        <strong>${formatNum(card.value)}</strong>
      </article>
    `
    )
    .join('');
}

function renderPagesByYear() {
  const rows = state.payload?.metrics?.avgPagesByYear || [];
  if (!rows.length) {
    pagesByYearChartEl.innerHTML = '<text x="16" y="40">No year/page data available.</text>';
    return;
  }

  const width = 940;
  const height = 360;
  const pad = { t: 20, r: 20, b: 40, l: 58 };
  const xs = rows.map((d) => d.year);
  const ys = rows.map((d) => d.mean || 0);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const maxY = Math.max(...ys, 1);

  const x = (v) => pad.l + ((v - minX) / Math.max(maxX - minX, 1)) * (width - pad.l - pad.r);
  const y = (v) => height - pad.b - (v / maxY) * (height - pad.t - pad.b);

  const points = rows.map((d) => `${x(d.year)},${y(d.mean || 0)}`).join(' ');

  const yTicks = Array.from({ length: 6 }, (_, i) => {
    const val = (maxY / 5) * i;
    return { val, y: y(val) };
  });

  pagesByYearChartEl.innerHTML = `
    ${yTicks
      .map(
        (tick) => `
      <line x1="${pad.l}" y1="${tick.y}" x2="${width - pad.r}" y2="${tick.y}" stroke="rgba(8,90,99,0.12)"/>
      <text class="axis" x="${pad.l - 8}" y="${tick.y + 4}" text-anchor="end">${formatNum(tick.val)}</text>
    `
      )
      .join('')}
    <polyline points="${points}" fill="none" stroke="#085a63" stroke-width="3" />
    ${rows
      .filter((_, i) => i % Math.ceil(rows.length / 12) === 0)
      .map((row) => `<text class="axis" x="${x(row.year)}" y="${height - 10}" text-anchor="middle">${row.year}</text>`)
      .join('')}
  `;
}

function renderDissertationsByYear() {
  const rows = state.payload?.metrics?.byYear || [];
  if (!rows.length) {
    dissertationsByYearChartEl.innerHTML = '<text x="16" y="40">No year data available.</text>';
    return;
  }

  const width = 940;
  const height = 360;
  const pad = { t: 20, r: 20, b: 40, l: 58 };
  const xs = rows.map((d) => d.year);
  const ys = rows.map((d) => d.count || 0);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const maxY = Math.max(...ys, 1);
  const barWidth = Math.max(4, ((width - pad.l - pad.r) / Math.max(rows.length, 1)) * 0.7);
  const barGap = (width - pad.l - pad.r) / Math.max(rows.length, 1);

  const x = (i) => pad.l + i * barGap + (barGap - barWidth) / 2;
  const y = (v) => height - pad.b - (v / maxY) * (height - pad.t - pad.b);

  const yTicks = Array.from({ length: 6 }, (_, i) => {
    const val = (maxY / 5) * i;
    return { val, y: y(val) };
  });

  dissertationsByYearChartEl.innerHTML = `
    ${yTicks
      .map(
        (tick) => `
      <line x1="${pad.l}" y1="${tick.y}" x2="${width - pad.r}" y2="${tick.y}" stroke="rgba(8,90,99,0.12)"/>
      <text class="axis" x="${pad.l - 8}" y="${tick.y + 4}" text-anchor="end">${formatNum(tick.val)}</text>
    `
      )
      .join('')}
    ${rows
      .map(
        (row, i) => `<rect x="${x(i)}" y="${y(row.count)}" width="${barWidth}" height="${height - pad.b - y(row.count)}" fill="var(--accent-2)" rx="2" />`
      )
      .join('')}
    ${rows
      .filter((_, i) => i % Math.ceil(rows.length / 12) === 0)
      .map((row, _, arr) => {
        const idx = rows.indexOf(row);
        return `<text class="axis" x="${x(idx) + barWidth / 2}" y="${height - 10}" text-anchor="middle">${row.year}</text>`;
      })
      .join('')}
  `;
}

function renderWordsByYear() {
  const rows = state.payload?.metrics?.byYear || [];
  if (!rows.length) {
    wordsByYearChartEl.innerHTML = '<text x="16" y="40">No year/word data available.</text>';
    return;
  }

  const width = 940;
  const height = 360;
  const pad = { t: 20, r: 20, b: 40, l: 58 };
  const xs = rows.map((d) => d.year);
  const ys = rows.map((d) => d.mean || 0);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const maxY = Math.max(...ys, 1);

  const x = (v) => pad.l + ((v - minX) / Math.max(maxX - minX, 1)) * (width - pad.l - pad.r);
  const y = (v) => height - pad.b - (v / maxY) * (height - pad.t - pad.b);

  const points = rows.map((d) => `${x(d.year)},${y(d.mean || 0)}`).join(' ');

  const yTicks = Array.from({ length: 6 }, (_, i) => {
    const val = (maxY / 5) * i;
    return { val, y: y(val) };
  });

  wordsByYearChartEl.innerHTML = `
    ${yTicks
      .map(
        (tick) => `
      <line x1="${pad.l}" y1="${tick.y}" x2="${width - pad.r}" y2="${tick.y}" stroke="rgba(8,90,99,0.12)"/>
      <text class="axis" x="${pad.l - 8}" y="${tick.y + 4}" text-anchor="end">${formatNum(tick.val)}</text>
    `
      )
      .join('')}
    <polyline points="${points}" fill="none" stroke="var(--accent-3)" stroke-width="3" />
    ${rows
      .filter((_, i) => i % Math.ceil(rows.length / 12) === 0)
      .map((row) => `<text class="axis" x="${x(row.year)}" y="${height - 10}" text-anchor="middle">${row.year}</text>`)
      .join('')}
  `;
}

function renderWordCloud() {
  const words = state.payload?.wordCloud || [];
  if (!words.length) {
    wordCloudEl.innerHTML = '<span>No theme terms available.</span>';
    themeResultsEl.innerHTML = '';
    return;
  }

  const max = Math.max(...words.map((w) => w.count), 1);
  wordCloudEl.innerHTML = words
    .map((word) => {
      const ratio = word.count / max;
      const size = 0.8 + ratio * 1.7;
      const hue = 190 - Math.round(ratio * 70);
      const active = state.selectedTheme && state.selectedTheme.toLowerCase() === word.term.toLowerCase();
      return `<button class="cloud-term${active ? ' active' : ''}" data-theme="${escapeHtml(word.term)}" style="font-size:${size}rem;color:hsl(${hue} 58% 28%)">${escapeHtml(word.term)}</button>`;
    })
    .join('');

  for (const node of wordCloudEl.querySelectorAll('.cloud-term[data-theme]')) {
    node.addEventListener('click', () => {
      state.selectedTheme = node.getAttribute('data-theme');
      renderWordCloud();
      openMatchesModal(`Theme: ${state.selectedTheme}`, docsForTheme(state.selectedTheme));
    });
  }

  themeResultsEl.innerHTML = '<p>Select a theme to view tagged dissertations.</p>';
}

function renderSubjectBars() {
  const byConcept = state.payload?.metrics?.byConcept || [];
  if (!byConcept.length) {
    subjectBarsEl.innerHTML = '<p style="color:var(--ink-soft);font-family:var(--sans);font-size:0.85rem">No concept length data available.</p>';
    return;
  }
  const maxMean = Math.max(...byConcept.map((s) => s.weightedMean || 0), 1);

  subjectBarsEl.innerHTML = byConcept
    .slice(0, 14)
    .map((entry) => {
      const widthPct = ((entry.weightedMean || 0) / maxMean) * 100;
      const label = `${entry.concept} (n=${formatNum(entry.docCount)})`;
      return `
        <div class="bar-row">
          <span class="bar-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span>
          <div class="bar-track"><div class="bar-fill" style="width:${widthPct}%"></div></div>
          <span class="bar-value">${formatNum(entry.weightedMean)}</span>
        </div>
      `;
    })
    .join('');
}

function renderPageTrend() {
  const rows = state.payload?.metrics?.pageTrend || [];
  if (!rows.length) {
    pageTrendChartEl.innerHTML = '<text x="16" y="40">No page trend data available.</text>';
    return;
  }

  const width = 940;
  const height = 360;
  const pad = { t: 20, r: 20, b: 40, l: 58 };
  const xs = rows.map((d) => d.year);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const maxY = Math.max(...rows.map((d) => d.max), 1);

  const x = (v) => pad.l + ((v - minX) / Math.max(maxX - minX, 1)) * (width - pad.l - pad.r);
  const y = (v) => height - pad.b - (v / maxY) * (height - pad.t - pad.b);

  const bandTop = rows.map((d) => `${x(d.year)},${y(d.max)}`).join(' ');
  const bandBot = [...rows].reverse().map((d) => `${x(d.year)},${y(d.min)}`).join(' ');
  const bandPoints = `${bandTop} ${bandBot}`;

  const medianPoints = rows.map((d) => `${x(d.year)},${y(d.median)}`).join(' ');

  const yTicks = Array.from({ length: 6 }, (_, i) => {
    const val = (maxY / 5) * i;
    return { val, y: y(val) };
  });

  pageTrendChartEl.innerHTML = `
    ${yTicks
      .map(
        (tick) => `
      <line x1="${pad.l}" y1="${tick.y}" x2="${width - pad.r}" y2="${tick.y}" stroke="rgba(8,90,99,0.12)"/>
      <text class="axis" x="${pad.l - 8}" y="${tick.y + 4}" text-anchor="end">${formatNum(tick.val)}</text>
    `
      )
      .join('')}
    <polygon points="${bandPoints}" fill="rgba(8,90,99,0.12)" />
    <polyline points="${medianPoints}" fill="none" stroke="#085a63" stroke-width="3" />
    ${rows
      .filter((_, i) => i % Math.ceil(rows.length / 12) === 0)
      .map((row) => `<text class="axis" x="${x(row.year)}" y="${height - 10}" text-anchor="middle">${row.year}</text>`)
      .join('')}
  `;
}

function renderNgramCloud() {
  const words = state.payload?.ngramCloud || [];
  if (!words.length) {
    ngramCloudEl.innerHTML = '<span>No n-gram data available.</span>';
    return;
  }

  const max = Math.max(...words.map((w) => w.count), 1);
  ngramCloudEl.innerHTML = words
    .map((word) => {
      const ratio = word.count / max;
      const size = 0.8 + ratio * 1.4;
      const hue = 20 + Math.round(ratio * 40);
      return `<button class="cloud-term" data-ngram-term="${escapeHtml(word.term)}" style="font-size:${size}rem;color:hsl(${hue} 68% 35%)">${escapeHtml(word.term)} <sup style="font-size:0.6em;opacity:0.6">${word.count}</sup></button>`;
    })
    .join('');

  for (const node of ngramCloudEl.querySelectorAll('.cloud-term[data-ngram-term]')) {
    node.addEventListener('click', () => {
      const term = node.getAttribute('data-ngram-term');
      openMatchesModal(`Concept: ${term}`, docsForConceptTerm(term));
    });
  }
}

function renderMethodologies() {
  const items = state.payload?.methodologies || [];
  if (!items.length) {
    methodologyBarsEl.innerHTML = '<p style="color:var(--ink-soft);font-family:var(--sans);font-size:0.85rem">No methodology signals detected.</p>';
    return;
  }

  const maxCount = Math.max(...items.map((m) => m.count), 1);
  methodologyBarsEl.innerHTML = items
    .map((entry) => {
      const widthPct = (entry.count / maxCount) * 100;
      return `
        <div class="bar-row">
          <span class="bar-label" title="${escapeHtml(entry.methodology)}">${escapeHtml(entry.methodology)}</span>
          <div class="bar-track"><div class="bar-fill" style="width:${widthPct}%"></div></div>
          <button class="bar-value" data-methodology="${escapeHtml(entry.methodology)}">${formatNum(entry.count)}</button>
        </div>
      `;
    })
    .join('');

  for (const node of methodologyBarsEl.querySelectorAll('[data-methodology]')) {
    node.addEventListener('click', () => {
      const methodology = node.getAttribute('data-methodology');
      openMatchesModal(`Methodology: ${methodology}`, docsForMethodology(methodology));
    });
  }
}

function renderCooccurrence() {
  const pairs = state.payload?.termCooccurrence || [];
  if (!pairs.length) {
    cooccurrenceBarsEl.innerHTML = '<p style="color:var(--ink-soft);font-family:var(--sans);font-size:0.85rem">No co-occurring term pairs found.</p>';
    return;
  }

  const maxCount = Math.max(...pairs.map((p) => p.count), 1);
  cooccurrenceBarsEl.innerHTML = pairs
    .map((entry) => {
      const widthPct = (entry.count / maxCount) * 100;
      const label = `${entry.termA} + ${entry.termB}`;
      return `
        <div class="bar-row">
          <span class="bar-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span>
          <div class="bar-track"><div class="bar-fill" style="width:${widthPct}%"></div></div>
          <button class="bar-value" data-co-term-a="${escapeHtml(entry.termA)}" data-co-term-b="${escapeHtml(entry.termB)}">${formatNum(entry.count)}</button>
        </div>
      `;
    })
    .join('');

  for (const node of cooccurrenceBarsEl.querySelectorAll('[data-co-term-a][data-co-term-b]')) {
    node.addEventListener('click', () => {
      const termA = node.getAttribute('data-co-term-a');
      const termB = node.getAttribute('data-co-term-b');
      openMatchesModal(`Co-occurrence: ${termA} + ${termB}`, docsForCooccurrence(termA, termB));
    });
  }
}

function renderSupervisorHeatmap() {
  const data = state.payload?.supervisorNgramMatrix;
  if (!data || !data.supervisors.length || !data.ngrams.length) {
    supervisorHeatmapEl.innerHTML = '<p style="color:var(--ink-soft);font-family:var(--sans);font-size:0.85rem">No supervisor-term data available.</p>';
    return;
  }

  const maxVal = Math.max(...data.matrix.flat(), 1);

  const headerCells = data.ngrams
    .map((s) => `<th class="heatmap-header">${escapeHtml(s)}</th>`)
    .join('');

  const bodyRows = data.supervisors
    .map((sup, si) => {
      const cells = data.ngrams
        .map((concept, nj) => {
          const val = data.matrix[si][nj];
          const lightness = val > 0 ? 95 - Math.round((val / maxVal) * 65) : 97;
          const textColor = lightness < 55 ? '#fff' : 'var(--ink)';
          const content = val > 0
            ? `<button class="heatmap-cell-btn" data-heatmap-sup="${escapeHtml(sup)}" data-heatmap-concept="${escapeHtml(concept)}" style="color:${textColor}">${val}</button>`
            : '';
          return `<td class="heatmap-cell" style="background:hsl(190 58% ${lightness}%);color:${textColor}">${content}</td>`;
        })
        .join('');
      return `<tr><td class="heatmap-label" title="${escapeHtml(sup)}"><button class="supervisor-link" data-supervisor-name="${escapeHtml(sup)}">${escapeHtml(sup)}</button></td>${cells}</tr>`;
    })
    .join('');

  supervisorHeatmapEl.innerHTML = `
    <table class="heatmap-table">
      <thead><tr><th></th>${headerCells}</tr></thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;

  for (const node of supervisorHeatmapEl.querySelectorAll('[data-heatmap-sup][data-heatmap-concept]')) {
    node.addEventListener('click', () => {
      const sup = node.getAttribute('data-heatmap-sup');
      const concept = node.getAttribute('data-heatmap-concept');
      openMatchesModal(`Supervisor + Concept: ${sup} + ${concept}`, docsForSupervisorConcept(sup, concept));
    });
  }

  for (const btn of supervisorHeatmapEl.querySelectorAll('.supervisor-link[data-supervisor-name]')) {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openSupervisorProfile(btn.dataset.supervisorName);
    });
  }
}

// --- Citation Explorer ---

function renderCitationDocs() {
  let docs = state.payload?.documents || [];
  if (state.citationFilterText) {
    const q = state.citationFilterText.toLowerCase();
    docs = docs.filter((doc) =>
      (doc.title || '').toLowerCase().includes(q) ||
      (doc.author || '').toLowerCase().includes(q) ||
      String(doc.year || '').includes(q)
    );
  }

  docs = [...docs].sort((a, b) => (b.citationCount || 0) - (a.citationCount || 0));

  const withCitations = docs.filter((d) => d.citationCount > 0).length;
  const headerEl = citationDocsTableEl.closest('.panel')?.querySelector('h2');
  if (headerEl) headerEl.textContent = `Dissertations (${withCitations} with parsed citations)`;

  citationDocsTableEl.innerHTML = docs
    .map((doc) => {
      const count = doc.citationCount || 0;
      const active = doc.id === state.citationDocId ? ' active' : '';
      const greyed = count === 0 ? ' greyed' : '';
      return `
        <tr class="citation-doc-row${active}${greyed}" data-cite-doc-id="${escapeHtml(doc.id)}">
          <td>${escapeHtml(doc.title || '(Untitled)')}</td>
          <td>${escapeHtml(doc.author || '')}</td>
          <td>${doc.year || '-'}</td>
          <td>${formatNum(count)}</td>
        </tr>
      `;
    })
    .join('');

  for (const row of citationDocsTableEl.querySelectorAll('.citation-doc-row')) {
    row.addEventListener('click', () => {
      selectCitationDoc(row.dataset.citeDocId);
    });
  }
}

async function selectCitationDoc(docId) {
  const requestToken = ++state.citationRequestToken;
  state.citationDocId = docId;
  renderCitationDocs();

  citationListTitleEl.textContent = 'Works Cited';
  citationEntriesEl.innerHTML = '<p class="meta">Loading citations...</p>';

  try {
    const res = await fetch(`/api/documents/${encodeURIComponent(docId)}/citations`);
    if (requestToken !== state.citationRequestToken) return;
    if (!res.ok) {
      citationEntriesEl.innerHTML = '<p class="meta">Failed to load citations.</p>';
      return;
    }
    const data = await res.json();
    if (requestToken !== state.citationRequestToken) return;
    const citations = data.citations || [];
    if (!citations.length) {
      citationEntriesEl.innerHTML = '<p class="meta">No citations found.</p>';
      citationListTitleEl.textContent = 'Works Cited';
      return;
    }
    citationListTitleEl.textContent = `Works Cited (${citations.length})`;
    renderCitationList(citations);
  } catch {
    if (requestToken !== state.citationRequestToken) return;
    citationEntriesEl.innerHTML = '<p class="meta">Connection error.</p>';
  }
}

function catalogueBadge(citation) {
  if (citation.catalogue_hits == null) return '';
  if (citation.catalogue_hits > 0) {
    const label = `Found in UBC Library (${citation.catalogue_hits} hit${citation.catalogue_hits !== 1 ? 's' : ''})`;
    if (citation.catalogue_bib_id) {
      return `<a class="catalogue-badge held" href="https://webcat.library.ubc.ca/vwebv/holdingsInfo?bibId=${encodeURIComponent(citation.catalogue_bib_id)}" target="_blank" rel="noreferrer" title="${label}" onclick="event.stopPropagation()">UBC Library</a>`;
    }
    return `<span class="catalogue-badge held" title="${label}">UBC Library</span>`;
  }
  return `<button class="catalogue-badge summon-check-btn" data-citation-id="${escapeHtml(String(citation.id))}" title="Check UBC Summon for this item" onclick="event.stopPropagation()">Check Summon</button>`;
}

function renderCitationList(citations) {
  state.selectedCitationIds = new Set();
  const selectAllHtml = `<div class="citation-select-all"><label><input type="checkbox" id="selectAllCitations" /> Select all</label></div>`;
  citationEntriesEl.innerHTML = selectAllHtml + citations
    .map((c) => {
      const citeCount = Math.max(1, Number(c.total_docs) || 1);
      const badge = `<span class="citation-count">${formatNum(citeCount)}</span>`;
      return `<div class="citation-entry" title="${escapeHtml(c.citation_text)}" data-citation-id="${c.id}" data-citation-text="${escapeHtml(c.citation_text)}" data-citation-count="${citeCount}"><input type="checkbox" class="citation-entry-check" data-check-cite-id="${c.id}" />${escapeHtml(c.citation_text)}${badge}${catalogueBadge(c)}</div>`;
    })
    .join('');

  const selectAllCb = document.getElementById('selectAllCitations');
  selectAllCb.addEventListener('change', () => {
    for (const cb of citationEntriesEl.querySelectorAll('.citation-entry-check')) {
      cb.checked = selectAllCb.checked;
      const id = cb.dataset.checkCiteId;
      if (selectAllCb.checked) state.selectedCitationIds.add(id);
      else state.selectedCitationIds.delete(id);
    }
  });

  for (const cb of citationEntriesEl.querySelectorAll('.citation-entry-check')) {
    cb.addEventListener('change', (e) => {
      e.stopPropagation();
      const id = cb.dataset.checkCiteId;
      if (cb.checked) state.selectedCitationIds.add(id);
      else state.selectedCitationIds.delete(id);
      const allChecks = citationEntriesEl.querySelectorAll('.citation-entry-check');
      const allChecked = Array.from(allChecks).every((c) => c.checked);
      selectAllCb.checked = allChecked;
      selectAllCb.indeterminate = !allChecked && state.selectedCitationIds.size > 0;
    });
    cb.addEventListener('click', (e) => e.stopPropagation());
  }

  for (const entry of citationEntriesEl.querySelectorAll('.citation-entry[data-citation-id]')) {
    entry.addEventListener('click', (e) => {
      if (e.target.closest('.citation-entry-check')) return;
      showCitingDissertations(
        entry.dataset.citationId,
        entry.dataset.citationText,
        Number(entry.dataset.citationCount || 1)
      );
    });
  }

  for (const btn of citationEntriesEl.querySelectorAll('.summon-check-btn')) {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const citationId = btn.dataset.citationId;
      btn.textContent = 'Checking\u2026';
      btn.disabled = true;
      try {
        const res = await fetch(`/api/citations/${encodeURIComponent(citationId)}/summon-check`);
        const data = await res.json();
        const replacement = document.createElement(data.found ? 'a' : 'a');
        replacement.className = `catalogue-badge ${data.found ? 'held' : 'not-held'}`;
        replacement.target = '_blank';
        replacement.rel = 'noreferrer';
        replacement.addEventListener('click', (ev) => ev.stopPropagation());
        if (data.found) {
          replacement.href = data.link;
          replacement.title = data.title || 'Found in UBC Library via Summon';
          replacement.textContent = 'Found in Summon';
        } else {
          replacement.href = data.illUrl || 'https://ill-docdel.library.ubc.ca/home';
          replacement.title = 'Not found in UBC Library \u2014 request via Interlibrary Loan / Document Delivery';
          replacement.textContent = 'Not found \u2014 ILL/DocDel';
        }
        btn.replaceWith(replacement);
      } catch {
        btn.textContent = 'Check Summon';
        btn.disabled = false;
      }
    });
  }
}

async function showCitingDissertations(citationId, citationText, totalDocs = null) {
  try {
    const res = await fetch(`/api/citations/${encodeURIComponent(citationId)}/documents`);
    if (!res.ok) return;
    const data = await res.json();
    const docs = (data.documents || []).map((d) => ({
      id: d.id,
      title: d.title || '(Untitled)',
      author: d.author || 'Unknown',
    }));

    const shortText = citationText.length > 140 ? citationText.slice(0, 140) + '...' : citationText;
    const linkedCount = totalDocs === null ? docs.length : totalDocs;
    const list = docs.length
      ? `<div class="related-list">
          ${docs.map((d) => `
            <div class="related-item" data-cite-nav-id="${escapeHtml(d.id)}">
              <strong>${escapeHtml(d.title)}</strong>
              <p>${escapeHtml(d.author)}</p>
            </div>
          `).join('')}
        </div>`
      : '<p class="meta">No dissertations found for this citation.</p>';

    docDetailsEl.innerHTML = `
      <div class="meta">
        <p><strong>Citation:</strong> ${escapeHtml(shortText)}</p>
        <p><strong>Linked dissertations:</strong> ${formatNum(linkedCount)}</p>
        <p>${formatNum(docs.length)} dissertation(s) loaded from index</p>
      </div>
      ${list}
    `;

    for (const item of docDetailsEl.querySelectorAll('.related-item[data-cite-nav-id]')) {
      item.addEventListener('click', () => {
        const targetId = item.getAttribute('data-cite-nav-id');
        if (targetId) openRecord(targetId, 'citations');
      });
    }
    docModalOverlay.hidden = false;
  } catch {
    // silently fail
  }
}

// --- Export utilities ---

function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function sanitizeBibKey(text) {
  return String(text || 'unknown').replace(/[^a-zA-Z0-9]/g, '').slice(0, 20);
}

function generateBibTeX(docs) {
  return docs.map((doc) => {
    const key = `${sanitizeBibKey(doc.author)}${doc.year || ''}`;
    const lines = [`@phdthesis{${key},`];
    lines.push(`  author = {${(doc.author || 'Unknown').replace(/[{}]/g, '')}},`);
    lines.push(`  title = {${(doc.title || '').replace(/[{}]/g, '')}},`);
    if (doc.year) lines.push(`  year = {${doc.year}},`);
    lines.push(`  school = {University of British Columbia},`);
    if (doc.degree) lines.push(`  type = {${doc.degree.replace(/[{}]/g, '')}},`);
    if (doc.doi) lines.push(`  doi = {${doc.doi}},`);
    if (doc.uri) lines.push(`  url = {${doc.uri}},`);
    lines.push('}');
    return lines.join('\n');
  }).join('\n\n');
}

function generateRIS(docs) {
  return docs.map((doc) => {
    const lines = ['TY  - THES'];
    lines.push(`AU  - ${doc.author || 'Unknown'}`);
    lines.push(`TI  - ${doc.title || ''}`);
    if (doc.year) lines.push(`PY  - ${doc.year}`);
    lines.push('PB  - University of British Columbia');
    if (doc.degree) lines.push(`M3  - ${doc.degree}`);
    if (doc.doi) lines.push(`DO  - ${doc.doi}`);
    if (doc.uri) lines.push(`UR  - ${doc.uri}`);
    if (doc.abstract) lines.push(`AB  - ${doc.abstract.slice(0, 500)}`);
    lines.push('ER  - ');
    return lines.join('\n');
  }).join('\n');
}

function generateCitationBibTeX(citations) {
  return citations.map((text, i) => {
    const key = `cite${i + 1}`;
    return `@misc{${key},\n  note = {${text.replace(/[{}]/g, '')}}\n}`;
  }).join('\n\n');
}

function generateCitationRIS(citations) {
  return citations.map((text) => {
    return `TY  - GEN\nT1  - ${text}\nER  - `;
  }).join('\n');
}

// --- Foundational Works ---

async function loadFoundationalWorks() {
  if (!foundationalWorksListEl) return;
  foundationalWorksListEl.innerHTML = '<p class="meta">Loading...</p>';
  try {
    const res = await fetch('/api/citations/top?limit=50');
    if (!res.ok) {
      foundationalWorksListEl.innerHTML = '<p class="meta">Could not load foundational works.</p>';
      return;
    }
    const data = await res.json();
    renderFoundationalWorks(data.works || []);
  } catch {
    foundationalWorksListEl.innerHTML = '<p class="meta">Connection error.</p>';
  }
}

function renderFoundationalWorks(works) {
  if (!works.length) {
    foundationalWorksListEl.innerHTML = '<p class="meta">No works cited across multiple dissertations yet.</p>';
    return;
  }

  foundationalWorksListEl.innerHTML = works.map((w) => {
    const badge = w.catalogue_hits > 0
      ? (w.catalogue_bib_id
        ? `<a class="catalogue-badge held" href="https://webcat.library.ubc.ca/vwebv/holdingsInfo?bibId=${encodeURIComponent(w.catalogue_bib_id)}" target="_blank" rel="noreferrer" onclick="event.stopPropagation()">UBC Library</a>`
        : '<span class="catalogue-badge held">UBC Library</span>')
      : '';
    return `
      <div class="foundational-work-item" data-citation-id="${w.id}" data-citation-text="${escapeHtml(w.citation_text)}" data-citation-count="${w.doc_count}">
        <span class="fw-rank-badge">${formatNum(w.doc_count)}</span>
        <span class="fw-text">${escapeHtml(w.citation_text)}${badge}</span>
      </div>
    `;
  }).join('');

  for (const item of foundationalWorksListEl.querySelectorAll('.foundational-work-item')) {
    item.addEventListener('click', () => {
      showCitingDissertations(
        item.dataset.citationId,
        item.dataset.citationText,
        Number(item.dataset.citationCount || 1)
      );
    });
  }
}

// --- Concept Timeline ---

function renderConceptTimeline() {
  const data = state.payload?.conceptTimeline || [];
  if (!data.length || !conceptTimelineChartEl) {
    if (conceptTimelineChartEl) conceptTimelineChartEl.innerHTML = '<text x="16" y="40" class="axis">No concept timeline data available.</text>';
    if (conceptTimelineLegendEl) conceptTimelineLegendEl.innerHTML = '';
    return;
  }

  const width = 940;
  const height = 360;
  const pad = { t: 20, r: 20, b: 40, l: 58 };

  const allYears = new Set();
  let maxCount = 1;
  for (const series of data) {
    for (const pt of series.data) {
      allYears.add(pt.year);
      if (pt.count > maxCount) maxCount = pt.count;
    }
  }
  const years = Array.from(allYears).sort((a, b) => a - b);
  if (!years.length) {
    conceptTimelineChartEl.innerHTML = '<text x="16" y="40" class="axis">No year data.</text>';
    conceptTimelineLegendEl.innerHTML = '';
    return;
  }

  const minX = years[0];
  const maxX = years[years.length - 1];
  const x = (v) => pad.l + ((v - minX) / Math.max(maxX - minX, 1)) * (width - pad.l - pad.r);
  const y = (v) => height - pad.b - (v / maxCount) * (height - pad.t - pad.b);

  const yTicks = Array.from({ length: 6 }, (_, i) => {
    const val = (maxCount / 5) * i;
    return { val, y: y(val) };
  });

  const hueStep = 360 / data.length;
  const lines = data.map((series, idx) => {
    const hue = Math.round(idx * hueStep);
    const color = `hsl(${hue} 65% 45%)`;
    const points = series.data.map((pt) => `${x(pt.year)},${y(pt.count)}`).join(' ');
    return `<polyline points="${points}" fill="none" stroke="${color}" stroke-width="2.5" />`;
  }).join('');

  const xLabels = years
    .filter((_, i) => i % Math.ceil(years.length / 12) === 0)
    .map((yr) => `<text class="axis" x="${x(yr)}" y="${height - 10}" text-anchor="middle">${yr}</text>`)
    .join('');

  conceptTimelineChartEl.innerHTML = `
    ${yTicks.map((tick) => `
      <line x1="${pad.l}" y1="${tick.y}" x2="${width - pad.r}" y2="${tick.y}" stroke="rgba(8,90,99,0.12)"/>
      <text class="axis" x="${pad.l - 8}" y="${tick.y + 4}" text-anchor="end">${formatNum(tick.val)}</text>
    `).join('')}
    ${lines}
    ${xLabels}
  `;

  conceptTimelineLegendEl.innerHTML = data.map((series, idx) => {
    const hue = Math.round(idx * hueStep);
    const color = `hsl(${hue} 65% 45%)`;
    return `<span class="timeline-legend-item"><span class="timeline-legend-swatch" style="background:${color}"></span>${escapeHtml(series.concept)} (${series.totalDocs})</span>`;
  }).join('');
}

// --- Supervisor Profiles ---

function buildSupervisorProfile(name, docs) {
  const supervised = docs.filter((d) => (d.supervisors || []).some((s) => s === name));
  const years = supervised.map((d) => d.year).filter(Boolean).sort((a, b) => a - b);
  const conceptMap = new Map();
  const methMap = new Map();
  for (const doc of supervised) {
    for (const c of (doc.conceptTerms || [])) {
      conceptMap.set(c, (conceptMap.get(c) || 0) + 1);
    }
    for (const m of (doc.methodologies || [])) {
      methMap.set(m, (methMap.get(m) || 0) + 1);
    }
  }
  const topConcepts = Array.from(conceptMap.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 12)
    .map(([term, count]) => ({ term, count }));
  const methodologies = Array.from(methMap.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([methodology, count]) => ({ methodology, count }));

  return {
    name,
    count: supervised.length,
    yearRange: years.length ? `${years[0]}\u2013${years[years.length - 1]}` : '-',
    dissertations: supervised,
    topConcepts,
    methodologies
  };
}

function renderSupervisorProfile(profile) {
  docModalTitleEl.textContent = `Supervisor: ${profile.name}`;

  const concepts = profile.topConcepts.length
    ? profile.topConcepts.map((c) => `<span class="token concept">${escapeHtml(c.term)} (${c.count})</span>`).join('')
    : '<span class="token concept">No concepts</span>';

  const maxMeth = Math.max(...profile.methodologies.map((m) => m.count), 1);
  const methBars = profile.methodologies.length
    ? profile.methodologies.map((m) => {
        const widthPct = (m.count / maxMeth) * 100;
        return `
          <div class="bar-row">
            <span class="bar-label">${escapeHtml(m.methodology)}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${widthPct}%"></div></div>
            <span class="bar-value">${formatNum(m.count)}</span>
          </div>
        `;
      }).join('')
    : '<p class="meta">No methodology signals.</p>';

  const dissertationList = profile.dissertations.length
    ? profile.dissertations.map((doc) => `
        <div class="related-item" data-related-id="${escapeHtml(doc.id)}">
          <strong>${escapeHtml(doc.title || '(Untitled)')}</strong>
          <p>${escapeHtml(doc.author || 'Unknown')} &middot; ${doc.year || '-'} &middot; ${escapeHtml(doc.degree || '-')}</p>
        </div>
      `).join('')
    : '<p class="meta">No dissertations found.</p>';

  docDetailsEl.innerHTML = `
    <div class="meta">
      <p><strong>${escapeHtml(profile.name)}</strong></p>
      <p>${formatNum(profile.count)} dissertation(s) &middot; ${profile.yearRange}</p>
    </div>
    <div>
      <p class="detail-section-title">Top Concepts</p>
      <div class="token-list">${concepts}</div>
    </div>
    <div>
      <p class="detail-section-title">Methodologies</p>
      <div class="bars">${methBars}</div>
    </div>
    <div>
      <p class="detail-section-title">Supervised Dissertations</p>
      <div class="related-list">${dissertationList}</div>
    </div>
  `;

  for (const item of docDetailsEl.querySelectorAll('.related-item[data-related-id]')) {
    item.addEventListener('click', () => {
      const targetId = item.getAttribute('data-related-id');
      if (targetId) openRecord(targetId, 'records');
    });
  }
  docModalOverlay.hidden = false;
}

function openSupervisorProfile(name) {
  const docs = state.payload?.documents || [];
  const profile = buildSupervisorProfile(name, docs);
  renderSupervisorProfile(profile);
}

// --- Methodology-Concept Matrix ---

function docsForMethodologyConcept(methodology, concept) {
  const methNorm = String(methodology || '').toLowerCase();
  const conceptNorm = String(concept || '').toLowerCase();
  const docs = state.payload?.documents || [];
  return docs.filter((doc) => {
    const hasMeth = (doc.methodologies || []).some((m) => String(m || '').toLowerCase() === methNorm);
    if (!hasMeth) return false;
    const terms = new Set((doc.conceptTerms || []).map((t) => String(t || '').toLowerCase()));
    return terms.has(conceptNorm);
  });
}

function renderMethodologyConceptMatrix() {
  const data = state.payload?.methodologyConceptMatrix;
  if (!data || !data.methodologies.length || !data.concepts.length) {
    if (methodologyConceptHeatmapEl) methodologyConceptHeatmapEl.innerHTML = '<p style="color:var(--ink-soft);font-family:var(--sans);font-size:0.85rem">No methodology-concept data available.</p>';
    return;
  }

  const maxVal = Math.max(...data.matrix.flat(), 1);

  const headerCells = data.concepts
    .map((c) => `<th class="heatmap-header">${escapeHtml(c)}</th>`)
    .join('');

  const bodyRows = data.methodologies
    .map((meth, mi) => {
      const cells = data.concepts
        .map((concept, ci) => {
          const val = data.matrix[mi][ci];
          const lightness = val > 0 ? 95 - Math.round((val / maxVal) * 65) : 97;
          const textColor = lightness < 55 ? '#fff' : 'var(--ink)';
          const content = val > 0
            ? `<button class="heatmap-cell-btn" data-mc-meth="${escapeHtml(meth)}" data-mc-concept="${escapeHtml(concept)}" style="color:${textColor}">${val}</button>`
            : '';
          return `<td class="heatmap-cell" style="background:hsl(30 58% ${lightness}%);color:${textColor}">${content}</td>`;
        })
        .join('');
      return `<tr><td class="heatmap-label" title="${escapeHtml(meth)}">${escapeHtml(meth)}</td>${cells}</tr>`;
    })
    .join('');

  methodologyConceptHeatmapEl.innerHTML = `
    <table class="heatmap-table">
      <thead><tr><th></th>${headerCells}</tr></thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;

  for (const node of methodologyConceptHeatmapEl.querySelectorAll('[data-mc-meth][data-mc-concept]')) {
    node.addEventListener('click', () => {
      const meth = node.getAttribute('data-mc-meth');
      const concept = node.getAttribute('data-mc-concept');
      openMatchesModal(`${meth} + ${concept}`, docsForMethodologyConcept(meth, concept));
    });
  }
}

// --- Research Gaps ---

function renderResearchGaps() {
  const gaps = state.payload?.researchGaps || [];
  if (!gaps.length) {
    if (researchGapsListEl) researchGapsListEl.innerHTML = '<p style="color:var(--ink-soft);font-family:var(--sans);font-size:0.85rem">No research gap data available.</p>';
    return;
  }

  const maxScore = Math.max(...gaps.map((g) => g.gapScore), 1);

  researchGapsListEl.innerHTML = gaps.map((entry) => {
    const widthPct = (entry.gapScore / maxScore) * 100;
    const label = `${entry.conceptA} + ${entry.conceptB}`;
    return `
      <div class="bar-row">
        <span class="bar-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span>
        <div class="bar-track"><div class="bar-fill gap-fill" style="width:${widthPct}%"></div></div>
        <span class="bar-value">${formatNum(Math.round(entry.gapScore))}</span>
      </div>
    `;
  }).join('');
}

function renderAll() {
  renderDocuments();
  renderKpis();
  renderPagesByYear();
  renderDissertationsByYear();
  renderWordsByYear();
  renderWordCloud();
  renderSubjectBars();
  renderPageTrend();
  renderNgramCloud();
  renderMethodologies();
  renderCooccurrence();
  renderSupervisorHeatmap();
  renderConceptTimeline();
  renderMethodologyConceptMatrix();
  renderResearchGaps();
}

// --- Data loading ---

async function loadData() {
  if (state.loading) return;
  state.loading = true;
  showSpinner(true);
  setStatus('Loading records...');

  try {
    const res = await fetch('/api/metrics');
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.message || err.error || `Request failed with ${res.status}`);
    }

    state.payload = await res.json();
    state.selectedDocId = state.payload.documents?.[0]?.id || null;
    state.selectedTheme = null;
    state.sortKey = null;
    state.sortDir = 'asc';
    state.filterText = '';
    state.selectedDocIds = new Set();
    docFilterEl.value = '';
    renderAll();

    const time = new Date(state.payload.generatedAt).toLocaleString();
    setStatus(`${formatNum(state.payload.metrics.recordCount)} documents loaded. Generated ${time}.`);
  } catch (error) {
    setStatus(`Failed to load data: ${error.message}`, true);
  } finally {
    state.loading = false;
    showSpinner(false);
  }
}

// --- Event bindings ---

for (const btn of tabButtons) {
  btn.addEventListener('click', () => {
    setActiveTab(btn.dataset.tab);
  });
}

for (const btn of citationTabButtons) {
  btn.addEventListener('click', () => setActiveCitationTab(btn.dataset.citationTab));
}

// Sort headers
for (const th of docTheadRow.querySelectorAll('th.sortable')) {
  th.addEventListener('click', () => {
    const key = th.dataset.sortKey;
    if (state.sortKey === key) {
      state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      state.sortKey = key;
      state.sortDir = 'asc';
    }
    renderDocuments();
  });
}

// Filter input
docFilterEl.addEventListener('input', () => {
  state.filterText = docFilterEl.value.trim();
  renderDocuments();
});

citationDocFilterEl.addEventListener('input', () => {
  state.citationFilterText = citationDocFilterEl.value.trim();
  renderCitationDocs();
});

docModalCloseBtn.addEventListener('click', closeDocModal);
docModalOverlay.addEventListener('click', (e) => {
  if (e.target === docModalOverlay) closeDocModal();
});
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !docModalOverlay.hidden) closeDocModal();
});

// Keyboard navigation for document table
documentsTableEl.addEventListener('keydown', (e) => {
  if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
  e.preventDefault();
  const docs = state.payload?.documents || [];
  if (!docs.length) return;
  const currentIndex = docs.findIndex((d) => d.id === state.selectedDocId);
  let nextIndex = currentIndex;
  if (e.key === 'ArrowDown') nextIndex = Math.min(currentIndex + 1, docs.length - 1);
  if (e.key === 'ArrowUp') nextIndex = Math.max(currentIndex - 1, 0);
  if (nextIndex !== currentIndex) {
    openRecord(docs[nextIndex].id, 'records');
    const row = documentsTableEl.querySelector(`[data-doc-id="${CSS.escape(docs[nextIndex].id)}"]`);
    row?.scrollIntoView({ block: 'nearest' });
  }
});

// Export buttons
exportBibTeXBtn.addEventListener('click', () => {
  const all = getFilteredSortedDocs();
  const docs = state.selectedDocIds.size > 0 ? all.filter((d) => state.selectedDocIds.has(d.id)) : all;
  if (!docs.length) return;
  downloadFile(generateBibTeX(docs), 'dissertations.bib', 'application/x-bibtex');
});

exportRISBtn.addEventListener('click', () => {
  const all = getFilteredSortedDocs();
  const docs = state.selectedDocIds.size > 0 ? all.filter((d) => state.selectedDocIds.has(d.id)) : all;
  if (!docs.length) return;
  downloadFile(generateRIS(docs), 'dissertations.ris', 'application/x-research-info-systems');
});

function getSelectedCitationTexts() {
  const entries = Array.from(citationEntriesEl.querySelectorAll('.citation-entry[data-citation-text]'));
  if (state.selectedCitationIds.size > 0) {
    return entries.filter((el) => state.selectedCitationIds.has(el.dataset.citationId)).map((el) => el.dataset.citationText);
  }
  return entries.map((el) => el.dataset.citationText);
}

exportCitationBibTeXBtn.addEventListener('click', () => {
  const texts = getSelectedCitationTexts();
  if (!texts.length) return;
  downloadFile(generateCitationBibTeX(texts), 'citations.bib', 'application/x-bibtex');
});

exportCitationRISBtn.addEventListener('click', () => {
  const texts = getSelectedCitationTexts();
  if (!texts.length) return;
  downloadFile(generateCitationRIS(texts), 'citations.ris', 'application/x-research-info-systems');
});

// Select-all docs checkbox
selectAllDocsEl.addEventListener('change', () => {
  const visibleChecks = documentsTableEl.querySelectorAll('.doc-row-check');
  for (const cb of visibleChecks) {
    const id = cb.dataset.checkId;
    if (selectAllDocsEl.checked) {
      cb.checked = true;
      state.selectedDocIds.add(id);
    } else {
      cb.checked = false;
      state.selectedDocIds.delete(id);
    }
  }
});

// Staggered reveal animation
for (const [idx, node] of document.querySelectorAll('.reveal').entries()) {
  node.style.animationDelay = `${idx * 70}ms`;
}

// Initial load
loadData();
