/*
 * Les 34 communes de la Martinique (liste officielle INSEE), pour
 * l'autocomplétion / liste déroulante du champ "Nom de la commune".
 * La saisie libre reste possible : ce n'est qu'une suggestion, pas une
 * contrainte — la normalisation (ElecParsers.normalizeCommune) rapproche de
 * toute façon "Le Marigot" et "MARIGOT" au moment de la vérification de
 * cohérence avec les fichiers importés.
 */
const MARTINIQUE_COMMUNES = [
  "Ajoupa-Bouillon", "Les Anses-d'Arlet", 'Basse-Pointe', 'Bellefontaine',
  'Le Carbet', 'Case-Pilote', 'Le Diamant', 'Ducos', 'Fonds-Saint-Denis',
  'Fort-de-France', 'Le François', "Grand'Rivière", 'Le Gros-Morne',
  'Le Lamentin', 'Le Lorrain', 'Macouba', 'Le Marigot', 'Le Marin',
  'Le Morne-Rouge', 'Le Morne-Vert', 'Le Prêcheur', 'Rivière-Pilote',
  'Rivière-Salée', 'Le Robert', 'Saint-Esprit', 'Saint-Joseph', 'Saint-Pierre',
  'Sainte-Anne', 'Sainte-Luce', 'Sainte-Marie', 'Schoelcher', 'La Trinité',
  "Les Trois-Îlets", 'Le Vauclin',
];
