/*
 * Lecture des fichiers sources envoyés par les communes :
 *  - fichier CAP  (onglet unique, entêtes détectées automatiquement)
 *  - fichier CCP / CST (onglets "contrat droit public" et "Contrat droit privé")
 *
 * Les lignes de titre/instructions des fichiers 2026 varient en nombre selon les
 * communes : on repère la ligne d'en-tête en cherchant la cellule "Collectivité..."
 * en colonne A plutôt que de fixer un numéro de ligne en dur.
 */

const ElecParsers = (() => {
  // Normalisation générale (accents, ligatures, apostrophes courbes, casse) —
  // utilisée pour les noms de personnes ET de communes. Ne touche jamais aux
  // espaces/tirets ni à un éventuel article : sur un nom de famille, retirer
  // un "LE"/"LA" en tête casserait des noms réels (ex. "LE GALL").
  function normalize(s) {
    return (s || '')
      .toString()
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/œ/gi, 'oe')
      .replace(/æ/gi, 'ae')
      .replace(/[’‘]/g, "'")
      .trim()
      .toUpperCase();
  }

  // Normalisation spécifique aux noms de commune, pour rapprocher les
  // variantes d'écriture qu'on rencontre en pratique : "Le Marigot" (nom
  // officiel INSEE, saisi dans l'appli) vs "MARIGOT" (valeur lue dans les
  // pré-listes, sans article). Ne pas appliquer à des noms de personnes.
  function normalizeCommune(s) {
    return normalize(s)
      .replace(/^(LE|LA|LES)\s+/, '')
      .replace(/^L'/, '')
      .replace(/[-\s]+/g, ' ')
      .trim();
  }

  function cellText(cell) {
    if (cell === null || cell === undefined) return '';
    if (cell instanceof Date) return cell;
    if (typeof cell === 'object' && cell.text) return cell.text;
    if (typeof cell === 'object' && cell.result !== undefined) return cell.result;
    return cell;
  }

  function findHeaderRow(worksheet, maxScan = 40) {
    for (let r = 1; r <= Math.min(maxScan, worksheet.rowCount); r++) {
      const a = normalize(cellText(worksheet.getRow(r).getCell(1).value));
      if (a.startsWith('COLLECTIVITE')) return r;
    }
    return null;
  }

  function readRows(worksheet, headerRow, columnCount) {
    const rows = [];
    for (let r = headerRow + 1; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const values = [];
      let hasContent = false;
      for (let c = 1; c <= columnCount; c++) {
        const v = cellText(row.getCell(c).value);
        if (v !== '' && v !== null && v !== undefined) hasContent = true;
        values.push(v);
      }
      if (!hasContent) continue;
      // A row is only real "data" if at least name or first name is present.
      rows.push({ rowNumber: r, values });
    }
    return rows;
  }

  function toDateOnly(v) {
    if (!v) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return v.toString().trim();
  }

  function splitCollectivite(raw) {
    const c = (raw || '').toString().trim().replace(/\s+/g, ' ');
    const up = normalize(c);
    const suffixes = ['CCAS', 'CE', 'COMMUNE'];
    for (const suf of suffixes) {
      if (up.endsWith(' ' + suf) || up === suf) {
        const commune = c.slice(0, c.length - suf.length).trim();
        return { commune, entite: suf, collectivite: c };
      }
    }
    return { commune: c, entite: 'AUTRE', collectivite: c };
  }

  // --- Fichier CAP -----------------------------------------------------

  async function parseCapFile(arrayBuffer, fileName) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);
    const ws = wb.worksheets[0];
    const headerRow = findHeaderRow(ws);
    if (headerRow === null) {
      throw new Error(
        `Impossible de trouver la ligne d'en-tête ("Collectivité...") dans ${fileName}.`
      );
    }
    // Colonnes attendues : Collectivité, Sexe, Nom usage, Nom famille, Prénom,
    // Date naissance, Département naissance, Code DEFI, Catégorie, Grade, Observations
    const rows = readRows(ws, headerRow, 11);
    const records = [];
    const warnings = [];
    for (const { rowNumber, values } of rows) {
      const [collectivite, sexe, nomUsage, nomFamille, prenom, dateNaissance,
        departement, codeDefi, categorie, grade, observation] = values;
      if (!nomFamille && !prenom) continue;
      const split = splitCollectivite(collectivite);
      const cat = (categorie || '').toString().trim().toUpperCase();
      const rec = {
        source: { fichier: fileName, onglet: ws.name, ligne: rowNumber, type: 'CAP' },
        collectivite: split.collectivite,
        commune: split.commune,
        entite: split.entite,
        sexe: (sexe || '').toString().trim(),
        nomUsage: (nomUsage || '').toString().trim(),
        nomFamille: (nomFamille || '').toString().trim(),
        prenom: (prenom || '').toString().trim(),
        dateNaissance: toDateOnly(dateNaissance),
        departementNaissance: (departement || '').toString().trim(),
        codeDefi: (codeDefi || '').toString().trim(),
        categorie: ['A', 'B', 'C'].includes(cat) ? cat : '',
        grade: (grade || '').toString().trim(),
        observation: (observation || '').toString().trim(),
        typeContrat: '', dateDebutContrat: '', dateFinContrat: '',
      };
      if (!rec.categorie) {
        warnings.push({
          type: 'categorie_absente',
          ref: `${fileName} L${rowNumber}`,
          message: `Catégorie CAP absente ou non reconnue (valeur lue : "${categorie || ''}")`,
          record: rec,
        });
      }
      records.push(rec);
    }
    return { records, warnings };
  }

  // --- Fichier CCP / CST -------------------------------------------------

  async function parseCcpCstFile(arrayBuffer, fileName) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);

    const findSheet = (needle) =>
      wb.worksheets.find((s) => normalize(s.name).includes(normalize(needle)));

    const publicSheet = findSheet('DROIT PUBLIC');
    const privateSheet = findSheet('DROIT PRIVE');

    const parseSheet = async (ws, kind) => {
      if (!ws) return { records: [], warnings: [], present: false };
      const headerRow = findHeaderRow(ws);
      if (headerRow === null) return { records: [], warnings: [], present: true, empty: true };
      // Collectivité, Sexe, Nom usage, Nom famille, Prénom, Date naissance,
      // Département naissance, Code DEFI, Catégorie, Grade, Type contrat,
      // Date début, Date fin, Observations
      const rows = readRows(ws, headerRow, 14);
      const records = [];
      for (const { rowNumber, values } of rows) {
        const [collectivite, sexe, nomUsage, nomFamille, prenom, dateNaissance,
          departement, codeDefi, categorie, grade, typeContrat, dateDebut, dateFin,
          observation] = values;
        if (!nomFamille && !prenom) continue;
        const split = splitCollectivite(collectivite);
        const cat = (categorie || '').toString().trim().toUpperCase();
        records.push({
          source: { fichier: fileName, onglet: ws.name, ligne: rowNumber, type: kind },
          collectivite: split.collectivite,
          commune: split.commune,
          entite: split.entite,
          sexe: (sexe || '').toString().trim(),
          nomUsage: (nomUsage || '').toString().trim(),
          nomFamille: (nomFamille || '').toString().trim(),
          prenom: (prenom || '').toString().trim(),
          dateNaissance: toDateOnly(dateNaissance),
          departementNaissance: (departement || '').toString().trim(),
          codeDefi: (codeDefi || '').toString().trim(),
          categorie: ['A', 'B', 'C'].includes(cat) ? cat : (cat || ''),
          grade: (grade || '').toString().trim(),
          typeContrat: (typeContrat || '').toString().trim(),
          dateDebutContrat: toDateOnly(dateDebut),
          dateFinContrat: toDateOnly(dateFin),
          observation: (observation || '').toString().trim(),
        });
      }
      return { records, warnings: [], present: true, empty: rows.length === 0 };
    };

    const pub = await parseSheet(publicSheet, 'CCP_PUBLIC');
    const priv = await parseSheet(privateSheet, 'CCP_PRIVE');

    const warnings = [];
    if (!pub.present) {
      warnings.push({
        type: 'onglet_absent',
        ref: fileName,
        message: `Onglet "contrat droit public" introuvable dans ${fileName}.`,
      });
    } else if (pub.empty) {
      warnings.push({
        type: 'onglet_vide',
        ref: fileName,
        message: `Onglet "contrat droit public" vide dans ${fileName} (0 ligne reçue — à confirmer, ce n'est pas une preuve d'absence d'agents).`,
      });
    }
    if (priv.present && priv.empty) {
      warnings.push({
        type: 'onglet_vide',
        ref: fileName,
        message: `Onglet "Contrat droit privé" vide dans ${fileName} (0 ligne reçue — à confirmer, ce n'est pas une preuve d'absence d'agents).`,
      });
    }

    // Détection automatique "avec CST" : l'onglet droit privé existe et a une
    // structure exploitable (même vide, sa présence signale une collectivité
    // rattachée au CST du CDG). Confirmable/à corriger manuellement dans l'UI.
    const avecCst = !!priv.present;

    return {
      recordsPublic: pub.records,
      recordsPrive: priv.records,
      warnings,
      avecCstDetecte: avecCst,
    };
  }

  return { parseCapFile, parseCcpCstFile, splitCollectivite, normalize, normalizeCommune };
})();
