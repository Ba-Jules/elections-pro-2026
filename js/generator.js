/*
 * Génération des fichiers .xlsx, fidèle à la structure des modèles 2022
 * (titre en Arial 18 gras centré fusionné B:G, mention légale éditable,
 * compteur "N inscrits", en-têtes de colonnes identiques, ligne de
 * signature finale), avec deux colonnes de traçabilité optionnelles.
 */

const ElecGenerator = (() => {
  const TITLES = {
    CAP: 'ELECTION AUX COMMISSIONS ADMINISTRATIVES PARITAIRES',
    CCP: 'ELECTION A LA COMMISSION CONSULTATIVE PARITAIRE',
    CST: 'ELECTION AU COMITE SOCIAL TERRITORIAL',
  };

  const COL_WIDTHS = [10, 36.11, 23, 19.33, 20.66, 19.11, 68.11, 26, 34];
  const BASE_HEADERS = [
    "N° d'ordre", 'Collectivité de gestion', "Nom d'usage", 'Nom de famille',
    'Prénom', 'Catégorie statutaire', 'Emploi - Grade',
  ];
  const TRACE_HEADERS = ['Référence source', 'Observation'];

  const MOIS_FR = [
    'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
    'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE',
  ];

  function formatDateScrutin(iso) {
    if (!iso) return '';
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    if (!m) return iso;
    const [, y, mo, d] = m;
    return `${parseInt(d, 10)} ${MOIS_FR[parseInt(mo, 10) - 1]} ${y}`;
  }

  function formatDateForFilename(iso) {
    const m = iso && /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    if (m) return `${m[3]}.${m[2]}.${m[1]}`;
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()}`;
  }

  function titleFont() {
    return { name: 'Arial', size: 18, bold: true };
  }
  function headerFont() {
    return { name: 'Calibri', size: 11, bold: true };
  }
  function bodyFont() {
    return { name: 'Calibri', size: 11 };
  }
  const CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };

  function mergeAndWrite(ws, range, value, font, alignment) {
    ws.mergeCells(range);
    const first = range.split(':')[0];
    const cell = ws.getCell(first);
    cell.value = value;
    cell.font = font;
    cell.alignment = alignment;
  }

  function buildListSheet(workbook, sheetName, { scrutinKey, categoryLabel, dateScrutin,
    mentionLegale, lieuSignature, dateGeneration, records, includeTrace }) {
    const ws = workbook.addWorksheet(sheetName);
    const headers = includeTrace ? [...BASE_HEADERS, ...TRACE_HEADERS] : BASE_HEADERS;
    const lastCol = String.fromCharCode(64 + headers.length); // 'G' or 'I'

    COL_WIDTHS.slice(0, headers.length).forEach((w, i) => {
      ws.getColumn(i + 1).width = w;
    });

    mergeAndWrite(ws, `B1:${lastCol}1`, TITLES[scrutinKey], titleFont(), CENTER);
    mergeAndWrite(ws, `B2:${lastCol}2`, `SCRUTIN DU ${formatDateScrutin(dateScrutin) || '[date à définir]'}`, titleFont(), CENTER);
    mergeAndWrite(ws, `B4:${lastCol}4`, 'LISTE ELECTORALE', titleFont(), CENTER);
    if (categoryLabel) {
      ws.getCell('A7').value = categoryLabel;
      ws.getCell('A7').font = { name: 'Calibri', size: 11, bold: true };
    }
    mergeAndWrite(
      ws, `A7:${lastCol}7`,
      categoryLabel ? categoryLabel : (mentionLegale || ''),
      { name: 'Calibri', size: 10, italic: !categoryLabel },
      { horizontal: categoryLabel ? 'left' : 'left', vertical: 'middle', wrapText: true }
    );
    if (categoryLabel) {
      ws.getCell('A7').value = categoryLabel;
    }
    ws.getRow(7).height = categoryLabel ? 18 : 40;

    ws.getCell(`${lastCol}8`).value = `${records.length} inscrits`;
    ws.getCell(`${lastCol}8`).font = { name: 'Calibri', size: 11, italic: true };
    ws.getCell(`${lastCol}8`).alignment = { horizontal: 'right' };

    const headerRowIdx = 9;
    const headerRow = ws.getRow(headerRowIdx);
    headers.forEach((h, i) => {
      const c = headerRow.getCell(i + 1);
      c.value = h;
      c.font = headerFont();
      c.alignment = CENTER;
      c.border = { bottom: { style: 'thin' } };
    });

    records.forEach((rec, idx) => {
      const r = ws.getRow(headerRowIdx + 1 + idx);
      const values = [
        idx + 1, rec.collectivite, rec.nomUsage, rec.nomFamille, rec.prenom,
        `Catégorie ${rec.categorie || ''}`.trim(), rec.grade,
      ];
      if (includeTrace) {
        values.push(`${rec.source.fichier} > ${rec.source.onglet} L${rec.source.ligne}`);
        values.push(rec.observation || '');
      }
      values.forEach((v, i) => {
        const c = r.getCell(i + 1);
        c.value = v;
        c.font = bodyFont();
        c.alignment = { horizontal: i === 6 || (includeTrace && i === 8) ? 'left' : 'center', vertical: 'middle' };
      });
    });

    const sigRowIdx = headerRowIdx + records.length + 2;
    mergeAndWrite(
      ws, `A${sigRowIdx}:${lastCol}${sigRowIdx}`,
      `Fait à ${lieuSignature || '[lieu à définir]'}, le ${dateGeneration}`,
      { name: 'Calibri', size: 11, italic: true },
      { horizontal: 'right' }
    );

    return ws;
  }

  async function buildSingleListWorkbook(fileTitle, sheetLabel, opts) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Consolidation listes électorales 2026 (outil local)';
    wb.created = new Date();
    buildListSheet(wb, sheetLabel, opts);
    return wb;
  }

  async function workbookToBlob(workbook) {
    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  const FILE_SPECS = [
    { key: 'capA', scrutinKey: 'CAP', sheet: 'CAP A', categoryLabel: 'CATEGORIE A', name: 'Liste CAP A compilée' },
    { key: 'capB', scrutinKey: 'CAP', sheet: 'CAP B', categoryLabel: 'CATEGORIE B', name: 'Liste CAP B compilée' },
    { key: 'capC', scrutinKey: 'CAP', sheet: 'CAP C', categoryLabel: 'CATEGORIE C', name: 'Liste CAP C compilée' },
    { key: 'ccp', scrutinKey: 'CCP', sheet: 'CCP', categoryLabel: null, name: 'Liste CCP compilée' },
    { key: 'cst', scrutinKey: 'CST', sheet: 'CST', categoryLabel: null, name: 'Liste CST compilée' },
  ];

  async function generateFiveFiles(compiled, settings) {
    const dateGeneration = new Date().toLocaleDateString('fr-FR');
    const files = [];
    for (const spec of FILE_SPECS) {
      const wb = await buildSingleListWorkbook(spec.name, spec.sheet, {
        scrutinKey: spec.scrutinKey,
        categoryLabel: spec.categoryLabel,
        dateScrutin: settings.dateScrutin,
        mentionLegale: settings.mentionLegale,
        lieuSignature: settings.lieuSignature,
        dateGeneration,
        records: compiled[spec.key],
        includeTrace: settings.includeTrace,
      });
      const blob = await workbookToBlob(wb);
      const filename = `${spec.name} - ${formatDateForFilename(settings.dateScrutin)}.xlsx`;
      files.push({ filename, blob });
    }
    return files;
  }

  async function generateExtracts(compiled, settings) {
    const dateGeneration = new Date().toLocaleDateString('fr-FR');
    const byCommune = new Map();
    const listDefs = [
      ['capA', 'CAP', 'CATEGORIE A', 'CAP A'],
      ['capB', 'CAP', 'CATEGORIE B', 'CAP B'],
      ['capC', 'CAP', 'CATEGORIE C', 'CAP C'],
      ['ccp', 'CCP', null, 'CCP'],
      ['cst', 'CST', null, 'CST'],
    ];
    for (const [key, scrutinKey, categoryLabel, label] of listDefs) {
      for (const rec of compiled[key]) {
        const commune = rec.commune || 'INCONNU';
        const entite = rec.entite || 'AUTRE';
        const mapKey = `${commune}||${entite}||${label}`;
        if (!byCommune.has(mapKey)) byCommune.set(mapKey, { commune, entite, scrutinKey, categoryLabel, label, records: [] });
        byCommune.get(mapKey).records.push(rec);
      }
    }
    const files = [];
    for (const grp of byCommune.values()) {
      const wb = await buildSingleListWorkbook(
        `${grp.commune} ${grp.entite} – ${grp.label}`,
        grp.label,
        {
          scrutinKey: grp.scrutinKey,
          categoryLabel: grp.categoryLabel,
          dateScrutin: settings.dateScrutin,
          mentionLegale: settings.mentionLegale,
          lieuSignature: settings.lieuSignature,
          dateGeneration,
          records: grp.records,
          includeTrace: settings.includeTrace,
        }
      );
      const blob = await workbookToBlob(wb);
      const filename = `${grp.commune} - ${grp.entite} - ${grp.label} - ${formatDateForFilename(settings.dateScrutin)}.xlsx`;
      files.push({ filename, blob });
    }
    return files;
  }

  async function generateExtractsZip(compiled, settings) {
    const files = await generateExtracts(compiled, settings);
    const zip = new JSZip();
    for (const f of files) zip.file(f.filename, f.blob);
    const content = await zip.generateAsync({ type: 'blob' });
    return { blob: content, count: files.length };
  }

  async function generateControlReport(state, compiled, settings) {
    const wb = new ExcelJS.Workbook();
    const synth = wb.addWorksheet('Synthèse');
    synth.columns = [
      { header: 'Commune', key: 'commune', width: 26 },
      { header: 'Entité', key: 'entite', width: 12 },
      { header: 'CAP A', key: 'capA', width: 8 },
      { header: 'CAP B', key: 'capB', width: 8 },
      { header: 'CAP C', key: 'capC', width: 8 },
      { header: 'CCP', key: 'ccp', width: 8 },
      { header: 'CST', key: 'cst', width: 8 },
    ];
    synth.getRow(1).font = { bold: true };
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
    for (const g of groups.values()) synth.addRow(g);

    const points = wb.addWorksheet('Points à examiner');
    points.columns = [
      { header: 'Type', key: 'type', width: 22 },
      { header: 'Référence', key: 'ref', width: 40 },
      { header: 'Détail', key: 'message', width: 90 },
    ];
    points.getRow(1).font = { bold: true };
    for (const a of compiled.anomalies) points.addRow({ type: a.type, ref: a.ref, message: a.message });

    const regles = wb.addWorksheet('Règles appliquées');
    regles.columns = [{ header: 'Réglage', key: 'k', width: 30 }, { header: 'Valeur', key: 'v', width: 80 }];
    regles.getRow(1).font = { bold: true };
    const cstRule = ElecRules.CST_RULES[settings.regleCst] || ElecRules.CST_RULES[ElecRules.DEFAULT_CST_RULE];
    regles.addRow({ k: 'Règle CST', v: cstRule.label });
    regles.addRow({ k: 'Date du scrutin', v: settings.dateScrutin || '(non définie)' });
    regles.addRow({ k: 'Lieu de signature', v: settings.lieuSignature || '(non défini)' });
    regles.addRow({ k: 'Colonnes de traçabilité', v: settings.includeTrace ? 'Incluses (H: Référence source, I: Observation)' : 'Masquées (strictement 7 colonnes)' });
    regles.addRow({ k: 'Communes importées', v: Object.keys(state.communes).join(', ') });

    return await workbookToBlob(wb);
  }

  return { generateFiveFiles, generateExtracts, generateExtractsZip, generateControlReport, triggerDownload };
})();
