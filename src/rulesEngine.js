// Moteur de règles - Étiquettes de prélèvement PET L30, L31, L32, L33, L34
// Basé sur QUAL/IT-606/VGZ rev15 — format 50cL uniquement

export const QUAND = [
  {
    id: "premier_po",
    label: "1er PO du jour",
    lignes: [
      { etiquette: "Analyses Labo", nature: 8, arome: 3 },
      { etiquette: "Hydrothèque", nature: 3, arome: 3 },
    ],
  },
  {
    id: "reprise_arret",
    label: "Reprise après arrêt > 4h",
    lignes: [{ etiquette: "St-10.000", nature: 4, arome: null }],
  },
  {
    id: "reprise_equipe_matin",
    label: "Redémarrage équipe matin < 4h",
    lignes: [
      { etiquette: "Analyses Labo", nature: 4, arome: null },
      { etiquette: "Hydrothèque", nature: 3, arome: null },
    ],
  },
  {
    id: "prelevement_3h_4h",
    label: "Prélèvement entre 3h et 4h",
    lignes: [
      { etiquette: "Analyses Labo", nature: 4, arome: null },
      { etiquette: "Hydrothèque", nature: 3, arome: null },
    ],
  },
  {
    id: "changement_article",
    label: "Changement d'article",
    lignes: [],
  },
];

export const POTYPES = [
  { id: "nature", label: "Nature" },
  { id: "arome", label: "Arôme" },
  { id: "finesbulles", label: "Fines bulles" },
];

// Combinaisons valides pour un changement d'article (PO actuel -> PO suivant).
// Les combinaisons absentes de cette table (Arôme->Arôme, Fines bulles->Fines bulles, etc.)
// n'existent pas dans le tableau qualité.
export const COMBOS = {
  "nature>nature": [{ etiquette: "Analyses Labo", qty: 4 }, { etiquette: "Hydrothèque", qty: 3 }],
  "arome>nature": [{ etiquette: "Analyses Labo", qty: 10 }, { etiquette: "Hydrothèque", qty: 3 }],
  "nature>arome": [{ etiquette: "Analyses Labo", qty: 3 }, { etiquette: "Hydrothèque", qty: 3 }],
  "nature>finesbulles": [{ etiquette: "Analyses Labo", qty: 5 }, { etiquette: "Hydrothèque", qty: 3 }],
  "finesbulles>nature": [{ etiquette: "Analyses Labo", qty: 5 }, { etiquette: "Hydrothèque", qty: 3 }],
};

export const USA_ETIQUETTE = "Étiquette Labo Noter US";
export const USA_QTY = 3;

/**
 * Calcule les étiquettes à imprimer pour une situation donnée.
 * @returns {{ resultats: Array<{etiquette: string, quantite: number}>, comboInconnu: boolean }}
 */
export function calculer({ quandId, poActuel, poSuivant, premierPoFlavor, usa }) {
  const lignesBrutes = [];
  const quand = QUAND.find((q) => q.id === quandId);
  if (!quand) return { resultats: [], comboInconnu: false };

  if (quandId === "premier_po") {
    for (const l of quand.lignes) {
      const val = premierPoFlavor === "arome" ? l.arome : l.nature;
      if (val) lignesBrutes.push({ etiquette: l.etiquette, quantite: val });
    }
  } else if (quandId !== "changement_article") {
    for (const l of quand.lignes) {
      if (l.nature) lignesBrutes.push({ etiquette: l.etiquette, quantite: l.nature });
    }
  }

  let comboInconnu = false;
  if (quandId === "changement_article" && poActuel && poSuivant) {
    const combo = COMBOS[`${poActuel}>${poSuivant}`];
    if (combo) {
      for (const l of combo) lignesBrutes.push({ etiquette: l.etiquette, quantite: l.qty });
    } else {
      comboInconnu = true;
    }
  }

  const map = new Map();
  for (const l of lignesBrutes) map.set(l.etiquette, (map.get(l.etiquette) || 0) + l.quantite);
  const resultats = Array.from(map, ([etiquette, quantite]) => ({ etiquette, quantite }));

  if (usa) resultats.push({ etiquette: USA_ETIQUETTE, quantite: USA_QTY });

  return { resultats, comboInconnu };
}
