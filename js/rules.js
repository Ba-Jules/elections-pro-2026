/*
 * Règles métier de compilation des listes électorales.
 * Toute hypothèse encore ambiguë entre les deux mails de JF est exposée comme un
 * réglage explicite (voir `settings.reglecst`) plutôt que tranchée en silence.
 */

const ElecRules = (() => {
  const { normalize, normalizeCommune } = ElecParsers;

  function dedupeKey(rec) {
    return [
      normalize(rec.nomFamille),
      normalize(rec.prenom),
      rec.dateNaissance || '',
      normalize(rec.commune),
    ].join('|');
  }

  function sourceRef(rec) {
    return `${rec.source.fichier} > ${rec.source.onglet} L${rec.source.ligne}`;
  }

  /**
   * Fusionne un lot d'enregistrements nouvellement importés dans l'état
   * cumulatif existant. Ne supprime jamais rien automatiquement : les
   * doublons probables sont signalés, pas fusionnés d'office.
   */
  function mergeIntoState(state, newRecords, listType) {
    const bucket = state.records[listType]; // 'CAP' | 'CCP' | 'CST'
    const seen = new Map(bucket.map((r) => [dedupeKey(r) + '|' + sourceRef(r), r]));
    const warnings = [];
    for (const rec of newRecords) {
      const key = dedupeKey(rec) + '|' + sourceRef(rec);
      if (seen.has(key)) continue; // ré-import exact du même fichier/ligne : ignoré
      const nearDuplicate = bucket.find((r) => dedupeKey(r) === dedupeKey(rec));
      if (nearDuplicate) {
        warnings.push({
          type: 'identite_repetee',
          ref: sourceRef(rec),
          message: `Identité déjà présente dans la liste ${listType} (rapprochement nom+prénom+naissance+commune) : ${rec.nomFamille} ${rec.prenom} — voir aussi ${sourceRef(nearDuplicate)}. Aucune suppression automatique.`,
        });
      }
      bucket.push(rec);
      seen.set(key, rec);
    }
    return warnings;
  }

  function sortAlpha(records) {
    return [...records].sort((a, b) => {
      const an = normalize(a.nomUsage) || normalize(a.nomFamille);
      const bn = normalize(b.nomUsage) || normalize(b.nomFamille);
      if (an !== bn) return an < bn ? -1 : 1;
      const af = normalize(a.nomFamille);
      const bf = normalize(b.nomFamille);
      if (af !== bf) return af < bf ? -1 : 1;
      const ap = normalize(a.prenom);
      const bp = normalize(b.prenom);
      return ap < bp ? -1 : ap > bp ? 1 : 0;
    });
  }

  // Options de résolution proposées pour chaque type d'anomalie. Affichées en
  // interface (cases à cocher / listes déroulantes) plutôt que laissées en
  // simple texte : chaque décision (celle d'Amadou, ou rapportée de JF) est
  // appliquée immédiatement aux listes compilées.
  const RESOLUTION_OPTIONS = {
    rattachement_incoherent: [
      { value: 'exclude', label: 'Exclure des listes (par défaut)' },
      { value: 'include', label: 'Inclure quand même, sous la collectivité détectée dans le fichier' },
    ],
    categorie_absente: [
      { value: 'exclude', label: 'Exclure des listes CAP (par défaut)' },
      { value: 'force_A', label: 'Forcer catégorie A' },
      { value: 'force_B', label: 'Forcer catégorie B' },
      { value: 'force_C', label: 'Forcer catégorie C' },
    ],
    identite_repetee: [
      { value: 'keep_both', label: 'Conserver les deux lignes (par défaut)' },
      { value: 'remove_this', label: 'Supprimer cette ligne (doublon)' },
    ],
    contrat_expire: [
      { value: 'keep', label: 'Conserver dans la liste (par défaut)' },
      { value: 'exclude', label: 'Exclure de la liste' },
    ],
    onglet_vide: [
      { value: 'pending', label: 'À vérifier (par défaut)' },
      { value: 'acknowledged', label: 'Confirmé — aucune donnée manquante' },
    ],
    onglet_absent: [
      { value: 'pending', label: 'À vérifier (par défaut)' },
      { value: 'acknowledged', label: 'Confirmé' },
    ],
  };

  // Règles possibles pour constituer la liste CST, sous forme de registre
  // nommé plutôt que d'un if/else : ajouter une variante plus tard (si JF en
  // demande une nuance) ne touche qu'à cette table, ni à l'UI ni à
  // buildCompiledLists.
  const DEFAULT_CST_RULE = 'fusion_cap_ccp';
  const CST_RULES = {
    fusion_cap_ccp: {
      label: 'Fusion CAP + CCP (public + privé) — collectivités "avec CST" (règle du 1er mail)',
      buildSource: ({ CAP, CCP_PUBLIC, CCP_PRIVE, cstFilter, resolveCategory }) => [
        ...CAP.filter(cstFilter).map((r) => ({ ...r, categorie: resolveCategory(r) })),
        ...CCP_PUBLIC.filter(cstFilter),
        ...CCP_PRIVE.filter(cstFilter),
      ],
    },
    ccp_cst_file_only: {
      label: 'Deux onglets du fichier CCP/CST uniquement (formulation littérale du mail détaillé)',
      buildSource: ({ CCP_PUBLIC, CCP_PRIVE, cstFilter }) => [
        ...CCP_PUBLIC.filter(cstFilter),
        ...CCP_PRIVE.filter(cstFilter),
      ],
    },
  };

  function resolutionKey(type, ref) {
    return `${type}::${ref}`;
  }

  function getResolution(state, type, ref) {
    const r = state.resolutions && state.resolutions[resolutionKey(type, ref)];
    return (r && r.action) || RESOLUTION_OPTIONS[type][0].value;
  }

  function setResolution(state, type, ref, action) {
    if (!state.resolutions) state.resolutions = {};
    state.resolutions[resolutionKey(type, ref)] = { action };
  }

  /**
   * Reconstruit les 5 listes officielles + les anomalies à partir de l'état
   * cumulatif (toutes communes importées jusqu'ici), des réglages courants,
   * et des décisions déjà prises sur chaque anomalie (state.resolutions).
   */
  function buildCompiledLists(state, settings) {
    if (!state.resolutions) state.resolutions = {};
    const { CAP, CCP_PUBLIC, CCP_PRIVE } = {
      CAP: state.records.CAP,
      CCP_PUBLIC: state.records.CCP_PUBLIC,
      CCP_PRIVE: state.records.CCP_PRIVE,
    };

    const anomalies = [...state.anomalies];

    // Rattachement incohérent : le nom de la commune déclarée à l'import ne
    // correspond pas à la collectivité lue dans les lignes du fichier.
    // Résolvable ligne à ligne ou en bloc via RESOLUTION_OPTIONS.rattachement_incoherent.
    const communeMatches = (a, b) => {
      const na = normalizeCommune(a);
      const nb = normalizeCommune(b);
      return na.includes(nb) || nb.includes(na);
    };
    const isCoherent = (rec) => {
      if (!rec.importCommuneDeclaree) return true;
      if (communeMatches(rec.commune, rec.importCommuneDeclaree)) return true;
      return getResolution(state, 'rattachement_incoherent', sourceRef(rec)) === 'include';
    };

    const allRecords = [...CAP, ...CCP_PUBLIC, ...CCP_PRIVE];
    for (const rec of allRecords) {
      const naturallyCoherent = !rec.importCommuneDeclaree || communeMatches(rec.commune, rec.importCommuneDeclaree);
      if (naturallyCoherent) continue;
      anomalies.push({
        type: 'rattachement_incoherent',
        ref: sourceRef(rec),
        message: `Fichier déclaré pour "${rec.importCommuneDeclaree}" mais ligne rattachée à "${rec.collectivite}".`,
        record: rec,
        resolved: getResolution(state, 'rattachement_incoherent', sourceRef(rec)),
      });
    }

    // Suppressions ponctuelles décidées sur des doublons détectés.
    const isRemovedDuplicate = (rec) =>
      getResolution(state, 'identite_repetee', sourceRef(rec)) === 'remove_this';

    // Catégorie CAP absente : résolvable en forçant A/B/C, sinon exclue.
    const resolveCategory = (rec) => {
      if (rec.categorie) return rec.categorie;
      const action = getResolution(state, 'categorie_absente', sourceRef(rec));
      return action.startsWith('force_') ? action.slice(6) : '';
    };
    for (const r of CAP.filter((r) => isCoherent(r) && !r.categorie)) {
      anomalies.push({
        type: 'categorie_absente',
        ref: sourceRef(r),
        message: `Catégorie CAP absente pour ${r.nomFamille} ${r.prenom}.`,
        record: r,
        resolved: getResolution(state, 'categorie_absente', sourceRef(r)),
      });
    }

    const capBase = CAP.filter((r) => isCoherent(r) && !isRemovedDuplicate(r))
      .map((r) => ({ ...r, categorie: resolveCategory(r) }))
      .filter((r) => r.categorie);
    const capA = sortAlpha(capBase.filter((r) => r.categorie === 'A'));
    const capB = sortAlpha(capBase.filter((r) => r.categorie === 'B'));
    const capC = sortAlpha(capBase.filter((r) => r.categorie === 'C'));

    // Contrats finissant avant la date du scrutin : résolvable (conserver / exclure).
    const isExcludedExpired = (r) => {
      if (!settings.dateScrutin || !r.dateFinContrat) return false;
      const fin = new Date(r.dateFinContrat);
      const scrutinDate = new Date(settings.dateScrutin);
      if (isNaN(fin) || fin >= scrutinDate) return false;
      return getResolution(state, 'contrat_expire', sourceRef(r)) === 'exclude';
    };
    if (settings.dateScrutin) {
      const scrutinDate = new Date(settings.dateScrutin);
      for (const r of [...CCP_PUBLIC, ...CCP_PRIVE]) {
        if (r.dateFinContrat) {
          const fin = new Date(r.dateFinContrat);
          if (!isNaN(fin) && fin < scrutinDate) {
            anomalies.push({
              type: 'contrat_expire',
              ref: sourceRef(r),
              message: `Contrat de ${r.nomFamille} ${r.prenom} se termine le ${r.dateFinContrat}, avant la date du scrutin (${settings.dateScrutin}).`,
              record: r,
              resolved: getResolution(state, 'contrat_expire', sourceRef(r)),
            });
          }
        }
      }
    }

    const ccpOk = sortAlpha(
      CCP_PUBLIC.filter((r) => isCoherent(r) && !isRemovedDuplicate(r) && !isExcludedExpired(r))
    );

    // CST : réglage explicite (voir plan), par commune "avec CST" seulement.
    // La règle appliquée est choisie dans CST_RULES (voir plus haut) plutôt
    // que codée en dur ici, pour qu'une future variante s'ajoute sans toucher
    // à cette fonction.
    const avecCstCommunes = new Set(
      Object.entries(state.communes)
        .filter(([, c]) => c.avecCst)
        .map(([nom]) => normalizeCommune(nom))
    );
    const estAvecCst = (rec) => avecCstCommunes.has(normalizeCommune(rec.commune));
    const cstFilter = (r) => isCoherent(r) && estAvecCst(r) && !isRemovedDuplicate(r) && !isExcludedExpired(r);

    const cstRule = CST_RULES[settings.regleCst] || CST_RULES[DEFAULT_CST_RULE];
    const cstSource = cstRule.buildSource({ CAP, CCP_PUBLIC, CCP_PRIVE, cstFilter, resolveCategory });
    // dédoublonnage interne (un agent peut apparaître à la fois en CAP et en CCP)
    const cstMap = new Map();
    for (const r of cstSource) cstMap.set(dedupeKey(r) + '|' + sourceRef(r), r);
    const cst = sortAlpha([...cstMap.values()]);

    // Annote chaque anomalie (y compris celles déjà stockées telles quelles
    // dans state.anomalies) avec la décision actuellement appliquée, pour que
    // l'interface puisse pré-sélectionner la bonne option.
    const finalAnomalies = anomalies.map((a) => ({
      ...a,
      resolved: a.resolved || getResolution(state, a.type, a.ref),
    }));

    return {
      capA, capB, capC, ccp: ccpOk, cst,
      anomalies: finalAnomalies,
      counts: { capA: capA.length, capB: capB.length, capC: capC.length, ccp: ccpOk.length, cst: cst.length },
    };
  }

  return {
    dedupeKey, sourceRef, mergeIntoState, sortAlpha, buildCompiledLists,
    RESOLUTION_OPTIONS, resolutionKey, getResolution, setResolution,
    CST_RULES, DEFAULT_CST_RULE,
  };
})();
