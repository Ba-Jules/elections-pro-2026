/*
 * État cumulatif du projet : toutes les communes importées, tous les
 * enregistrements consolidés, les anomalies, les réglages. Sauvegardable
 * dans un fichier local (etat_projet.json) pour permettre l'import
 * "au fur et à mesure" d'une session à l'autre, et autosauvegardé dans
 * localStorage pour reprendre facilement dans le même navigateur.
 */

const ElecState = (() => {
  const LOCALSTORAGE_KEY = 'elections_pro_state_v1';

  function emptyState() {
    return {
      version: 1,
      communes: {}, // { "AJOUPA-BOUILLON": { avecCst: true, importsCap: [...], importsCcp: [...] } }
      records: { CAP: [], CCP_PUBLIC: [], CCP_PRIVE: [] },
      anomalies: [],
      resolutions: {}, // clé "type::référence" -> { action } — décisions appliquées aux anomalies
      settings: {
        dateScrutin: '',
        lieuSignature: '',
        mentionLegale:
          "Conformément aux articles 9 et 10 du décret n°89-229 du 17 avril 1989, les électeurs peuvent, le cas échéant, présenter au Président du Centre de Gestion des demandes d'inscription ou des réclamations contre les inscriptions ou omissions de la liste électorale (délais à adapter à l'échéance 2026).",
        regleCst: 'fusion_cap_ccp',
        includeTrace: true,
      },
      log: [],
    };
  }

  function addLog(state, message) {
    state.log.unshift({ ts: new Date().toISOString(), message });
    if (state.log.length > 300) state.log.length = 300;
  }

  function saveToLocalStorage(state) {
    try {
      localStorage.setItem(LOCALSTORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      // stockage indisponible (navigation privée, quota) : silencieux, l'export
      // fichier reste la voie de sauvegarde fiable.
    }
  }

  function loadFromLocalStorage() {
    try {
      const raw = localStorage.getItem(LOCALSTORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  function exportStateFile(state) {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `etat_projet - ${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  async function importStateFile(file) {
    const text = await file.text();
    const parsed = JSON.parse(text);
    if (!parsed || !parsed.records) throw new Error('Fichier état de projet invalide.');
    return parsed;
  }

  return { emptyState, addLog, saveToLocalStorage, loadFromLocalStorage, exportStateFile, importStateFile };
})();
