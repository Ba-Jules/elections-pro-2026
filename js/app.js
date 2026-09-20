/* Orchestration de l'interface. Aucune donnée ne quitte le navigateur :
 * tout se passe en mémoire / fichiers locaux, aucun appel réseau. */

(() => {
  let state = ElecState.loadFromLocalStorage() || ElecState.emptyState();
  let compiled = null; // dernier résultat de compilation
  let blockSeq = 0;
  const blocks = new Map(); // id -> { commune, avecCstManual, capParsed, ccpParsed }

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  function toast(message, kind = 'info') {
    const zone = $('#toastZone');
    const el = document.createElement('div');
    el.className = `toast toast--${kind}`;
    el.textContent = message;
    zone.appendChild(el);
    requestAnimationFrame(() => el.classList.add('toast--visible'));
    setTimeout(() => {
      el.classList.remove('toast--visible');
      setTimeout(() => el.remove(), 300);
    }, 5000);
  }

  function log(message) {
    ElecState.addLog(state, message);
    renderLog();
  }

  function persist() {
    ElecState.saveToLocalStorage(state);
  }

  // ---------------------------------------------------------------- Réglages

  function populateStaticLists() {
    // Options de la règle CST : générées depuis ElecRules.CST_RULES, pas
    // codées en dur, pour qu'une future variante n'exige qu'une entrée dans
    // ce registre (voir rules.js).
    const select = $('#regleCst');
    select.innerHTML = Object.entries(ElecRules.CST_RULES)
      .map(([value, rule]) => `<option value="${value}">${rule.label}</option>`)
      .join('');

    // Autocomplétion du nom de commune : les 34 communes de Martinique,
    // en saisie libre malgré tout (une commune hors liste reste possible).
    $('#communesDatalist').innerHTML = MARTINIQUE_COMMUNES
      .map((c) => `<option value="${c}"></option>`)
      .join('');
  }

  let settingsListenersBound = false;
  function bindSettings() {
    const s = state.settings;
    $('#projectName').value = state.projectName || '';
    $('#dateScrutin').value = s.dateScrutin || '';
    $('#lieuSignature').value = s.lieuSignature || '';
    $('#mentionLegale').value = s.mentionLegale || '';
    $('#regleCst').value = s.regleCst || ElecRules.DEFAULT_CST_RULE;
    $('#includeTrace').checked = !!s.includeTrace;

    if (settingsListenersBound) return;
    settingsListenersBound = true;
    $('#projectName').addEventListener('input', (e) => { state.projectName = e.target.value; persist(); });
    $('#dateScrutin').addEventListener('change', (e) => { state.settings.dateScrutin = e.target.value; persist(); recompute(); });
    $('#lieuSignature').addEventListener('input', (e) => { state.settings.lieuSignature = e.target.value; persist(); });
    $('#mentionLegale').addEventListener('input', (e) => { state.settings.mentionLegale = e.target.value; persist(); });
    $('#regleCst').addEventListener('change', (e) => { state.settings.regleCst = e.target.value; persist(); recompute(); });
    $('#includeTrace').addEventListener('change', (e) => { state.settings.includeTrace = e.target.checked; persist(); });
  }

  // ------------------------------------------------------------ Import blocs

  function createCommuneBlock() {
    const id = `bloc-${++blockSeq}`;
    blocks.set(id, { commune: '', avecCstManual: null, capParsed: null, ccpParsed: null, pendingParses: 0 });

    const wrap = document.createElement('div');
    wrap.className = 'commune-block';
    wrap.id = id;
    wrap.innerHTML = `
      <div class="commune-block__head">
        <input type="text" class="commune-name" list="communesDatalist" placeholder="Nom de la commune (ex : Le Marigot)" autocomplete="off" />
        <button class="btn btn--ghost btn--small remove-block" title="Retirer ce bloc"><svg class="icon"><use href="#icon-x"/></svg></button>
      </div>
      <label class="avec-cst-row">
        <input type="checkbox" class="avec-cst-checkbox" />
        <span>Collectivité avec CST (détecté automatiquement à l'import du fichier CCP/CST — modifiable)</span>
      </label>
      <div class="dropzones">
        <div class="dropzone" data-kind="cap">
          <svg class="icon"><use href="#icon-import"/></svg>
          <div class="dropzone__label">Fichier CAP</div>
          <div class="dropzone__status">Glissez le fichier ici ou cliquez</div>
          <input type="file" accept=".xlsx" class="file-input" data-kind="cap" hidden />
        </div>
        <div class="dropzone" data-kind="ccp">
          <svg class="icon"><use href="#icon-import"/></svg>
          <div class="dropzone__label">Fichier CCP / CST</div>
          <div class="dropzone__status">Glissez le fichier ici ou cliquez</div>
          <input type="file" accept=".xlsx" class="file-input" data-kind="ccp" hidden />
        </div>
      </div>
      <div class="commune-block__warnings"></div>
      <button class="btn btn--tonal btn--small add-to-consolidation" disabled>Ajouter à la consolidation</button>
    `;
    $('#communeBlocks').appendChild(wrap);
    wireBlock(wrap, id);
    return wrap;
  }

  function wireBlock(wrap, id) {
    const info = blocks.get(id);

    wrap.querySelector('.commune-name').addEventListener('input', (e) => {
      info.commune = e.target.value.trim();
    });
    wrap.querySelector('.remove-block').addEventListener('click', () => {
      blocks.delete(id);
      wrap.remove();
    });
    wrap.querySelector('.avec-cst-checkbox').addEventListener('change', (e) => {
      info.avecCstManual = e.target.checked;
    });

    $$('.dropzone', wrap).forEach((zone) => {
      const kind = zone.dataset.kind;
      const input = zone.querySelector('.file-input');
      zone.addEventListener('click', () => input.click());
      zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dropzone--over'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('dropzone--over'));
      zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dropzone--over');
        if (e.dataTransfer.files[0]) handleFileForBlock(id, wrap, kind, e.dataTransfer.files[0]);
      });
      input.addEventListener('change', (e) => {
        if (e.target.files[0]) handleFileForBlock(id, wrap, kind, e.target.files[0]);
      });
    });

    wrap.querySelector('.add-to-consolidation').addEventListener('click', () => addBlockToConsolidation(id, wrap));
  }

  function updateAddButtonState(wrap, info) {
    const btn = wrap.querySelector('.add-to-consolidation');
    if (info.pendingParses > 0) {
      btn.disabled = true;
      btn.textContent = 'Lecture des fichiers en cours…';
    } else {
      btn.disabled = !(info.capParsed || info.ccpParsed);
      btn.textContent = 'Ajouter à la consolidation';
    }
  }

  async function handleFileForBlock(id, wrap, kind, file) {
    const info = blocks.get(id);
    const zone = wrap.querySelector(`.dropzone[data-kind="${kind}"]`);
    const statusEl = zone.querySelector('.dropzone__status');
    statusEl.textContent = 'Lecture en cours…';
    zone.classList.add('dropzone--busy');
    info.pendingParses++;
    updateAddButtonState(wrap, info);
    try {
      const buffer = await file.arrayBuffer();
      if (kind === 'cap') {
        const result = await ElecParsers.parseCapFile(buffer, file.name);
        info.capParsed = { file, ...result };
        statusEl.textContent = `${file.name} — ${result.records.length} ligne(s)`;
      } else {
        const result = await ElecParsers.parseCcpCstFile(buffer, file.name);
        info.ccpParsed = { file, ...result };
        statusEl.textContent = `${file.name} — ${result.recordsPublic.length} droit public / ${result.recordsPrive.length} droit privé`;
        if (info.avecCstManual === null) {
          info.avecCstManual = result.avecCstDetecte;
          wrap.querySelector('.avec-cst-checkbox').checked = result.avecCstDetecte;
        }
      }
      zone.classList.remove('dropzone--busy');
      zone.classList.add('dropzone--ok');
      renderBlockWarnings(wrap, info);
    } catch (err) {
      zone.classList.remove('dropzone--busy');
      zone.classList.add('dropzone--error');
      statusEl.textContent = `Erreur : ${err.message}`;
      toast(`Impossible de lire ${file.name} : ${err.message}`, 'error');
    } finally {
      info.pendingParses--;
      updateAddButtonState(wrap, info);
    }
  }

  function renderBlockWarnings(wrap, info) {
    const box = wrap.querySelector('.commune-block__warnings');
    const warnings = [
      ...(info.capParsed ? info.capParsed.warnings : []),
      ...(info.ccpParsed ? info.ccpParsed.warnings : []),
    ];
    if (!warnings.length) { box.innerHTML = ''; return; }
    box.innerHTML = `<div class="warn-list">${warnings
      .map((w) => `<div class="warn-item"><svg class="icon"><use href="#icon-alert"/></svg>${escapeHtml(w.message)}</div>`)
      .join('')}</div>`;
  }

  function addBlockToConsolidation(id, wrap) {
    const info = blocks.get(id);
    if (info.pendingParses > 0) {
      toast('Lecture des fichiers encore en cours, merci de patienter.', 'error');
      return;
    }
    const commune = wrap.querySelector('.commune-name').value.trim();
    if (!commune) { toast('Merci de renseigner le nom de la commune.', 'error'); return; }
    if (!info.capParsed && !info.ccpParsed) { toast('Aucun fichier importé pour cette commune.', 'error'); return; }

    const avecCst = wrap.querySelector('.avec-cst-checkbox').checked;
    state.communes[commune] = { avecCst };

    const tag = (recs) => recs.map((r) => ({ ...r, importCommuneDeclaree: commune }));

    let addedWarnings = [];
    if (info.capParsed) {
      addedWarnings = addedWarnings.concat(
        ElecRules.mergeIntoState(state, tag(info.capParsed.records), 'CAP')
      );
      // Les avertissements "catégorie absente" du parseur ne sont pas conservés ici :
      // ElecRules.buildCompiledLists les recalcule à chaque compilation (et les
      // rend résolvables), un doublon fixe ferait doublon avec cette liste vivante.
    }
    if (info.ccpParsed) {
      addedWarnings = addedWarnings.concat(
        ElecRules.mergeIntoState(state, tag(info.ccpParsed.recordsPublic), 'CCP_PUBLIC')
      );
      addedWarnings = addedWarnings.concat(
        ElecRules.mergeIntoState(state, tag(info.ccpParsed.recordsPrive), 'CCP_PRIVE')
      );
      state.anomalies.push(...info.ccpParsed.warnings);
    }
    state.anomalies.push(...addedWarnings);

    log(`Commune "${commune}" ajoutée à la consolidation (avec CST : ${avecCst ? 'oui' : 'non'}).`);
    toast(`${commune} ajoutée à la consolidation.`, 'success');
    persist();
    recompute();

    wrap.querySelector('.add-to-consolidation').disabled = true;
    wrap.querySelector('.add-to-consolidation').textContent = 'Ajoutée ✓';
    $$('.file-input', wrap).forEach((i) => (i.disabled = true));
  }

  // ------------------------------------------------------------- Compilation

  function recompute() {
    compiled = ElecRules.buildCompiledLists(state, state.settings);
    renderDashboard();
    renderAnomalies();
    renderExportSection();
    renderOverview();
  }

  function renderDashboard() {
    const box = $('#dashboard');
    const communeNames = Object.keys(state.communes);
    if (!communeNames.length) {
      box.innerHTML = '<div class="empty-state"><svg class="icon"><use href="#icon-check-circle"/></svg>Aucune commune consolidée pour le moment — déposez un premier fichier à gauche pour démarrer.</div>';
      return;
    }
    const groups = new Map();
    const bump = (rec, field) => {
      const k = `${rec.commune}||${rec.entite}`;
      if (!groups.has(k)) groups.set(k, { commune: rec.commune, entite: rec.entite, capA: 0, capB: 0, capC: 0, ccp: 0, cst: 0 });
      groups.get(k)[field]++;
    };
    compiled.capA.forEach((r) => bump(r, 'capA'));
    compiled.capB.forEach((r) => bump(r, 'capB'));
    compiled.capC.forEach((r) => bump(r, 'capC'));
    compiled.ccp.forEach((r) => bump(r, 'ccp'));
    compiled.cst.forEach((r) => bump(r, 'cst'));

    const rows = [...groups.values()]
      .sort((a, b) => a.commune.localeCompare(b.commune) || a.entite.localeCompare(b.entite))
      .map((g) => `<tr><td>${g.commune}</td><td>${g.entite}</td><td>${g.capA}</td><td>${g.capB}</td><td>${g.capC}</td><td>${g.ccp}</td><td>${g.cst}</td></tr>`)
      .join('');

    box.innerHTML = `
      <div class="table-scroll"><table class="table">
        <thead><tr><th>Commune</th><th>Entité</th><th>CAP A</th><th>CAP B</th><th>CAP C</th><th>CCP</th><th>CST</th></tr></thead>
        <tbody>${rows}</tbody>
        <tfoot><tr><td colspan="2">Total</td><td>${compiled.counts.capA}</td><td>${compiled.counts.capB}</td><td>${compiled.counts.capC}</td><td>${compiled.counts.ccp}</td><td>${compiled.counts.cst}</td></tr></tfoot>
      </table></div>`;
  }

  function anomalyGroupKey(a) {
    // Regroupe par type + fichier source, pour éviter une liste de centaines
    // de lignes quasi identiques quand une anomalie touche tout un fichier.
    const fileKey = (a.ref || '').replace(/\s*>.*$/, '').trim();
    return `${a.type}||${fileKey}`;
  }

  const openAnomalyGroups = new Set(); // conserve les groupes dépliés d'un rendu à l'autre

  function escapeAttr(s) {
    return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;');
  }
  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function resolutionSelectHtml(type, ref, current, cssClass) {
    const options = ElecRules.RESOLUTION_OPTIONS[type];
    if (!options) return '';
    const opts = options
      .map((o) => `<option value="${o.value}" ${o.value === current ? 'selected' : ''}>${escapeHtml(o.label)}</option>`)
      .join('');
    return `<select class="${cssClass}" data-type="${type}" data-ref="${escapeAttr(ref)}">${opts}</select>`;
  }

  function renderAnomalies() {
    const box = $('#anomalies');
    if (!compiled.anomalies.length) {
      box.innerHTML = '<div class="empty-state"><svg class="icon"><use href="#icon-check-circle"/></svg>Aucun point à examiner — toutes les lignes importées sont cohérentes.</div>';
      return;
    }
    const groups = new Map();
    for (const a of compiled.anomalies) {
      const k = anomalyGroupKey(a);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(a);
    }

    const anomalyBlocks = [...groups.entries()]
      .sort((a, b) => b[1].length - a[1].length)
      .map(([groupKey, items]) => {
        const first = items[0];
        const fileKey = (first.ref || '').replace(/\s*>.*$/, '').trim();
        const bulkHtml = items.length > 1
          ? `<select class="resolution-bulk" data-type="${first.type}" data-group="${escapeAttr(groupKey)}">
              <option value="">Appliquer à tout le groupe (${items.length} lignes)…</option>
              ${ElecRules.RESOLUTION_OPTIONS[first.type]
                .map((o) => `<option value="${o.value}">${escapeHtml(o.label)}</option>`)
                .join('')}
            </select>`
          : '';
        const rowsHtml = items
          .map((a) => `<tr><td>${escapeHtml(a.ref)}</td><td>${escapeHtml(a.message)}</td><td>${resolutionSelectHtml(a.type, a.ref, a.resolved, 'resolution-select')}</td></tr>`)
          .join('');
        const table = `<div class="table-scroll"><table class="table table--anomalies">
            <thead><tr><th>Référence</th><th>Détail</th><th>Décision</th></tr></thead>
            <tbody>${rowsHtml}</tbody>
          </table></div>`;
        if (items.length <= 5) {
          return `<div class="anomaly-group-block">
            <div class="anomaly-group-head"><strong>${labelForAnomalyType(first.type)} — ${escapeHtml(fileKey)}</strong></div>
            ${table}
          </div>`;
        }
        const isOpen = openAnomalyGroups.has(groupKey);
        return `<details class="anomaly-group" data-group="${escapeAttr(groupKey)}" ${isOpen ? 'open' : ''}>
          <summary>${labelForAnomalyType(first.type)} — ${escapeHtml(fileKey)} (${items.length} lignes)</summary>
          <div class="anomaly-group-head">${bulkHtml}</div>
          ${table}
        </details>`;
      });

    box.innerHTML = anomalyBlocks.join('');
  }

  function labelForAnomalyType(t) {
    return {
      rattachement_incoherent: 'Rattachement incohérent',
      categorie_absente: 'Catégorie CAP absente',
      identite_repetee: 'Identité répétée',
      contrat_expire: 'Contrat expirant avant le scrutin',
      onglet_vide: 'Onglet vide',
      onglet_absent: 'Onglet absent',
    }[t] || t;
  }

  function renderLog() {
    const box = $('#journal');
    box.innerHTML = state.log
      .slice(0, 50)
      .map((l) => `<div class="log-item"><span class="log-item__ts">${new Date(l.ts).toLocaleString('fr-FR')}</span> ${l.message}</div>`)
      .join('');
  }

  function renderExportSection() {
    const has = compiled && (compiled.capA.length || compiled.capB.length || compiled.capC.length || compiled.ccp.length || compiled.cst.length);
    $('#exportSection').classList.toggle('hidden', !has);
  }

  // ------------------------------------------------------------ Vue d'ensemble

  function renderOverview() {
    const box = $('#overviewPanel');
    const communeCount = Object.keys(state.communes).length;

    if (!communeCount) {
      box.innerHTML = `
        <div class="hero-empty">
          <div class="hero-empty__icon"><svg class="icon"><use href="#icon-building"/></svg></div>
          <div class="hero-empty__text">
            <h2>Aucune commune importée pour l'instant</h2>
            <p class="muted">Déposez le premier fichier CAP et le fichier CCP/CST d'une commune dans la section « Importation » pour démarrer la consolidation.</p>
          </div>
          <a href="#section-import" class="btn btn--primary"><svg class="icon"><use href="#icon-import"/></svg>Commencer l'import</a>
        </div>`;
      return;
    }

    const agentsRecenses = state.records.CAP.length + state.records.CCP_PUBLIC.length + state.records.CCP_PRIVE.length;
    const listesCompilees = compiled.counts.capA + compiled.counts.capB + compiled.counts.capC + compiled.counts.ccp + compiled.counts.cst;
    const anomalieCount = compiled.anomalies.length;
    const filesGenerated = !$('#exportSection').classList.contains('hidden');

    box.innerHTML = `
      <div class="overview-grid">
        <div class="kpi-card">
          <div class="kpi-icon"><svg class="icon"><use href="#icon-building"/></svg></div>
          <div><div class="kpi-value">${communeCount}</div><div class="kpi-label">Commune${communeCount > 1 ? 's' : ''} consolidée${communeCount > 1 ? 's' : ''}</div></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-icon"><svg class="icon"><use href="#icon-import"/></svg></div>
          <div><div class="kpi-value">${agentsRecenses}</div><div class="kpi-label">Lignes agents reçues</div><div class="kpi-sub">CAP + CCP/CST, avant filtrage</div></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-icon ${anomalieCount ? 'kpi-icon--warn' : ''}"><svg class="icon"><use href="#icon-alert"/></svg></div>
          <div><div class="kpi-value">${anomalieCount}</div><div class="kpi-label">Point${anomalieCount > 1 ? 's' : ''} à examiner</div></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-icon kpi-icon--champagne"><svg class="icon"><use href="#icon-results"/></svg></div>
          <div><div class="kpi-value">${listesCompilees}</div><div class="kpi-label">Agents dans les 5 listes</div><div class="kpi-sub">${filesGenerated ? 'Prêtes à télécharger' : 'Compilez pour générer les fichiers'}</div></div>
        </div>
      </div>
      <div class="overview-banner">
        <div class="overview-banner__text">
          <h2>${filesGenerated ? 'Les 5 listes sont à jour' : 'Prêt à compiler'}</h2>
          <p class="muted">${anomalieCount
            ? `${anomalieCount} point${anomalieCount > 1 ? 's' : ''} à examiner avant transmission — voir la section Contrôles.`
            : 'Aucune anomalie en attente sur les données actuellement importées.'}</p>
        </div>
        <a href="#section-consolidation" class="btn btn--primary"><svg class="icon"><use href="#icon-layers"/></svg>Aller à la consolidation</a>
      </div>`;
  }

  // ------------------------------------------------------------------ Export

  async function onCompileClick() {
    recompute();
    log('Compilation des 5 listes mise à jour.');
    toast('Listes recompilées.', 'success');
    document.getElementById('resultatsSection').scrollIntoView({ behavior: 'smooth' });
  }

  async function onDownloadAll() {
    const files = await ElecGenerator.generateFiveFiles(compiled, state.settings);
    files.forEach((f) => ElecGenerator.triggerDownload(f.blob, f.filename));
    log('Téléchargement des 5 fichiers compilés.');
  }

  async function onDownloadOne(key) {
    const files = await ElecGenerator.generateFiveFiles(compiled, state.settings);
    const idx = ['capA', 'capB', 'capC', 'ccp', 'cst'].indexOf(key);
    ElecGenerator.triggerDownload(files[idx].blob, files[idx].filename);
  }

  async function onGenerateExtracts() {
    const { blob, count } = await ElecGenerator.generateExtractsZip(compiled, state.settings);
    ElecGenerator.triggerDownload(blob, `Extraits par commune - ${new Date().toISOString().slice(0, 10)}.zip`);
    log(`${count} extraits générés et téléchargés (archive zip).`);
    toast(`${count} extraits générés.`, 'success');
  }

  async function onDownloadReport() {
    const blob = await ElecGenerator.generateControlReport(state, compiled, state.settings);
    ElecGenerator.triggerDownload(blob, `Rapport de contrôle - ${new Date().toISOString().slice(0, 10)}.xlsx`);
    log('Rapport de contrôle téléchargé.');
  }

  // -------------------------------------------------------------- État projet

  function onSaveState() {
    ElecState.exportStateFile(state);
    log('État du projet sauvegardé (fichier .json).');
  }

  async function onLoadState(file) {
    try {
      const loaded = await ElecState.importStateFile(file);
      state = loaded;
      persist();
      renderAll();
      log(`État du projet rechargé depuis ${file.name}.`);
      toast('Projet rechargé.', 'success');
    } catch (err) {
      toast(`Fichier d'état invalide : ${err.message}`, 'error');
    }
  }

  function onNewProject() {
    if (!confirm('Repartir sur un nouveau projet ? Les communes non sauvegardées seront perdues.')) return;
    state = ElecState.emptyState();
    compiled = null;
    persist();
    renderAll();
  }

  // ----------------------------------------------------------------- Wiring

  function renderAll() {
    bindSettings();
    $('#communeBlocks').innerHTML = '';
    blocks.clear();
    createCommuneBlock();
    renderLog();
    recompute();
  }

  function wireAnomalyResolutions() {
    const box = $('#anomalies');
    box.addEventListener('change', (e) => {
      if (e.target.matches('.resolution-select')) {
        const { type, ref } = e.target.dataset;
        ElecRules.setResolution(state, type, ref, e.target.value);
        persist();
        const label = e.target.options[e.target.selectedIndex].text;
        log(`Décision appliquée — ${labelForAnomalyType(type)} (${ref}) : ${label}.`);
        recompute();
      } else if (e.target.matches('.resolution-bulk')) {
        const { type, group } = e.target.dataset;
        const action = e.target.value;
        if (!action) return;
        const items = compiled.anomalies.filter((a) => anomalyGroupKey(a) === group);
        for (const a of items) ElecRules.setResolution(state, type, a.ref, action);
        persist();
        const label = e.target.options[e.target.selectedIndex].text;
        log(`Décision appliquée à ${items.length} ligne(s) — ${labelForAnomalyType(type)} : ${label}.`);
        toast(`Décision appliquée à ${items.length} ligne(s).`, 'success');
        recompute();
      }
    });
    // 'toggle' ne remonte pas (bubble) : on l'écoute en phase de capture.
    box.addEventListener('toggle', (e) => {
      if (e.target.matches && e.target.matches('details.anomaly-group')) {
        const key = e.target.dataset.group;
        if (e.target.open) openAnomalyGroups.add(key);
        else openAnomalyGroups.delete(key);
      }
    }, true);
  }

  // Met en évidence, dans la barre latérale et la barre mobile, la section
  // actuellement visible à l'écran (pure navigation, aucune donnée modifiée).
  function initScrollspy() {
    const sectionIds = [
      'section-overview', 'section-import', 'section-consolidation',
      'section-settings', 'section-anomalies', 'section-results', 'section-journal',
    ];
    const sections = sectionIds.map((id) => document.getElementById(id)).filter(Boolean);
    if (!sections.length || !('IntersectionObserver' in window)) return;
    const links = $$('.nav-link');
    const setActive = (id) => links.forEach((l) => l.classList.toggle('is-active', l.dataset.section === id));
    const observer = new IntersectionObserver((entries) => {
      const visible = entries.filter((e) => e.isIntersecting).sort((a, b) => b.intersectionRatio - a.intersectionRatio);
      if (visible.length) setActive(visible[0].target.id);
    }, { rootMargin: '-10% 0px -70% 0px', threshold: [0, 0.25, 0.5, 0.75, 1] });
    sections.forEach((s) => observer.observe(s));
    setActive(sections[0].id);
  }

  function init() {
    populateStaticLists();
    initScrollspy();
    $('#addCommuneBtn').addEventListener('click', createCommuneBlock);
    $('#compileBtn').addEventListener('click', onCompileClick);
    wireAnomalyResolutions();
    $('#downloadAllBtn').addEventListener('click', onDownloadAll);
    $$('.download-one').forEach((btn) => btn.addEventListener('click', () => onDownloadOne(btn.dataset.key)));
    $('#extractsBtn').addEventListener('click', onGenerateExtracts);
    $('#reportBtn').addEventListener('click', onDownloadReport);
    $('#saveStateBtn').addEventListener('click', onSaveState);
    $('#loadStateInput').addEventListener('change', (e) => {
      if (e.target.files[0]) onLoadState(e.target.files[0]);
    });
    $('#newProjectBtn').addEventListener('click', onNewProject);

    renderAll();
    if (Object.keys(state.communes).length) {
      toast('Projet précédent rechargé automatiquement depuis ce navigateur.', 'info');
    }
  }

  document.addEventListener('DOMContentLoaded', init);
})();
