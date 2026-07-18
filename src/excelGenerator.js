import ExcelJS from "exceljs";
import modeleUrl from "./assets/modele-etiquettes.xlsx?url";

// =============================================================================
// Générateur d'étiquettes de prélèvement — approche "modèle rempli".
//
// Le fichier src/assets/modele-etiquettes.xlsx est un export exact et non
// modifié de la feuille "Etiquettes Vierge V2" du classeur d'origine
// Etiquettes_prélèvements.xlsm (macros retirées, elles ne servaient qu'à
// faire ce que ce fichier fait déjà en JS). Hauteurs de lignes et polices
// viennent de ce fichier et ne sont jamais recalculées : ce module se
// contente d'écrire des valeurs dans des cellules déjà stylées. Seules
// exceptions, corrigées après coup car le modèle d'origine n'était pas fiable
// sur ces points précis : le format papier (voir PAPIER_A4), les marges de
// page et la largeur des colonnes (voir MARGES_PAGE et LARGEURS_COLONNES,
// valeurs reprises de la boîte de dialogue Page Setup et des propriétés de
// colonnes du classeur d'origine) et l'ajustement automatique à la page (voir
// plus bas) pour garantir une seule page quelle que soit la largeur
// réellement rendue par Excel.
//
// Structure de la feuille modèle (vérifiée sur le fichier d'origine) :
// grille de 7 bandes x 3 colonnes (A/B/C) = 21 étiquettes, une colonne =
// une étiquette entière. À partir de la ligne de titre de chaque bande
// ("PRELEVEMENTS", colonne A), le motif est toujours :
//   +0  titre "PRELEVEMENTS" (identique pour toutes les étiquettes)
//   +1  emplacement "type" en 10 pt (style "HYDROTHEQUE")
//   +2  emplacement "type" en 8 pt (style "ANALYSE LABO")
//   +3  "Date : ….      Ligne :"
//   +4  "Heure : ….      Equipe : "
//   +5  "Production : ….     "
//   +6  "Article : ….      "
//   puis 2 lignes ou plus d'espacement avant la bande suivante (l'espacement
//   réel entre bandes n'est PAS régulier dans le fichier d'origine — l'écart
//   entre la 4e et la 5e bande fait 11 lignes au lieu de 9 — donc la position
//   de chaque bande est détectée dynamiquement en cherchant "PRELEVEMENTS"
//   plutôt que calculée par un pas fixe).
// Selon le type d'étiquette, un seul des deux emplacements de type (+1 ou +2)
// reçoit du texte ; l'autre est vidé pour ne pas laisser le texte par défaut
// du modèle.
// =============================================================================

const COLONNES = ["A", "B", "C"];
const NB_BANDES_ATTENDU = 7;
const ETIQUETTES_PAR_PAGE = NB_BANDES_ATTENDU * COLONNES.length; // 21
const NOM_FEUILLE_MODELE = "Etiquettes Vierge V2";
const PAPIER_A4 = 9;

// Marges horizontales (onglet Margins de la boîte de dialogue Page Setup),
// reprises telles quelles depuis le classeur Excel d'origine. ExcelJS attend
// les marges en pouces, d'où la conversion cm -> pouces (cm / 2.54). Les
// marges verticales (top/bottom/header/footer) ne sont volontairement pas
// touchées : elles restent celles déjà présentes dans le modèle.
const cmVersPouces = (cm) => cm / 2.54;
const MARGES_HORIZONTALES = {
  left: cmVersPouces(1.4),
  right: cmVersPouces(0.4),
};

// Largeurs de colonnes du classeur d'origine (unités de largeur de colonne
// Excel), gardées distinctes entre A/B/C — pas d'uniformisation. Colonne A
// ajustée de 33.14 à 35.34 (demande explicite, confirmée avec Youssouf).
const LARGEURS_COLONNES_ORIGINE = { A: 35.34, B: 35.0, C: 31.0 };

// Écart supplémentaire demandé entre la colonne 1 (A) et la colonne 2 (B) :
// +2 mm, ajoutés à la largeur de la colonne A (ce qui repousse d'autant le
// début de la colonne B). Conversion mm -> largeur de colonne Excel avec la
// même correspondance que celle utilisée pour les largeurs ci-dessus
// (largeur_px = mm * 96/25.4 ; largeur_excel = largeur_px / 7).
const PX_PAR_MM = 96 / 25.4;
const GAP_SUPPLEMENTAIRE_A_B_MM = 2;
const LARGEURS_COLONNES = {
  ...LARGEURS_COLONNES_ORIGINE,
  A: LARGEURS_COLONNES_ORIGINE.A + (GAP_SUPPLEMENTAIRE_A_B_MM * PX_PAR_MM) / 7,
};

// 14 espaces à la place de l'heure quand elle est laissée vide, pour garder
// "Equipe" à la même position visuelle qu'un remplissage manuel du formulaire papier.
const ESPACEMENT_HEURE_VIDE = " ".repeat(14);

// Repère la ligne de départ de chaque bande d'étiquettes en cherchant le
// titre "PRELEVEMENTS" en colonne A, plutôt que de calculer sa position par
// un pas fixe (l'espacement entre bandes n'est pas régulier dans le modèle).
function detecterBasesDeBande(ws) {
  const bases = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    const valeur = ws.getCell(`A${r}`).value;
    if (typeof valeur === "string" && valeur.trim() === "PRELEVEMENTS") {
      bases.push(r);
    }
  }
  if (bases.length !== NB_BANDES_ATTENDU) {
    throw new Error(
      `Modèle Excel inattendu : ${bases.length} bande(s) d'étiquettes détectée(s) au lieu de ${NB_BANDES_ATTENDU}.`
    );
  }
  return bases;
}

// Pour chaque type d'étiquette produit par le moteur de règles : quel
// emplacement de type utiliser (1 = 10pt façon HYDROTHEQUE, 2 = 8pt façon
// ANALYSE LABO) et quel texte y écrire. Validé avec Youssouf : St-10.000
// reprend le style court (comme HYDROTHEQUE), Étiquette Labo Noter US
// s'imprime "PRODUCTION US" en style compact (comme ANALYSE LABO).
// "Hydrothèque" seul déroge à la taille de police par défaut de son
// emplacement (Arial 10) : on la force à Arial 8 via taillePolice, sans
// toucher St-10.000 qui reste sur le même emplacement à 10pt.
const TYPES_ETIQUETTE = {
  "Hydrothèque": { offset: 1, texte: " HYDROTHEQUE ", taillePolice: 8 },
  "Analyses Labo": { offset: 2, texte: "ANALYSE LABO" },
  "St-10.000": { offset: 1, texte: "ST-10.000" },
  "Étiquette Labo Noter US": { offset: 2, texte: "PRODUCTION US" },
};

// Remplace, dans un texte de cellule du modèle, la suite de points de
// pointillé ("…….." ou "......") par une valeur réelle, en laissant intact
// tout le reste du texte (préfixe, espacement, suffixe) tel qu'il est défini
// dans le modèle.
function injecterValeur(texteModele, valeur) {
  return texteModele.replace(/[….]+/, String(valeur ?? ""));
}

function chargerFeuilleModele(workbook) {
  const ws = workbook.getWorksheet(NOM_FEUILLE_MODELE);
  if (!ws) {
    throw new Error(
      `Feuille "${NOM_FEUILLE_MODELE}" introuvable dans le modèle Excel.`
    );
  }
  return ws;
}

function remplirEtiquette(ws, basesDeBande, etiquette, indexDansPage, infos) {
  const style = TYPES_ETIQUETTE[etiquette];
  if (!style) {
    throw new Error(`Type d'étiquette inconnu pour le modèle Excel : "${etiquette}"`);
  }

  const bande = Math.floor(indexDansPage / COLONNES.length);
  const col = COLONNES[indexDansPage % COLONNES.length];
  const base = basesDeBande[bande];

  const autreOffset = style.offset === 1 ? 2 : 1;
  const celluleType = ws.getCell(`${col}${base + style.offset}`);
  celluleType.value = style.texte;
  if (style.taillePolice) {
    // Toutes les cellules d'un même emplacement (offset 1 ou 2, sur les 7
    // bandes x 3 colonnes) partagent le même objet de style tant qu'aucune
    // n'a été stylée individuellement : assigner uniquement `.font`
    // mute cet objet partagé et change la taille sur TOUTES ces cellules
    // (y compris St-10.000, qui doit rester à 10pt). Il faut donc réassigner
    // `.style` en entier pour détacher cette cellule du style partagé.
    celluleType.style = {
      ...celluleType.style,
      font: { ...celluleType.font, size: style.taillePolice },
    };
  }
  ws.getCell(`${col}${base + autreOffset}`).value = "";

  const dateAffichee = infos.date
    ? new Date(infos.date + "T00:00:00").toLocaleDateString("fr-FR")
    : "";

  const celluleDateLigne = ws.getCell(`${col}${base + 3}`);
  celluleDateLigne.value =
    injecterValeur(celluleDateLigne.value, dateAffichee) + " " + (infos.ligne || "");

  const celluleHeureEquipe = ws.getCell(`${col}${base + 4}`);
  const valeurHeure = infos.heure || ESPACEMENT_HEURE_VIDE;
  celluleHeureEquipe.value = injecterValeur(celluleHeureEquipe.value, valeurHeure) + (infos.equipe || "");

  const celluleProduction = ws.getCell(`${col}${base + 5}`);
  celluleProduction.value = injecterValeur(celluleProduction.value, infos.production);

  const celluleArticle = ws.getCell(`${col}${base + 6}`);
  celluleArticle.value = injecterValeur(celluleArticle.value, infos.article);
}

// Vide entièrement un emplacement d'étiquette non utilisé (au-delà du nombre
// réellement demandé), pour ne pas laisser le texte d'exemple du modèle sur
// une étiquette qui ne sera pas remplie.
function viderEtiquette(ws, basesDeBande, indexDansPage) {
  const bande = Math.floor(indexDansPage / COLONNES.length);
  const col = COLONNES[indexDansPage % COLONNES.length];
  const base = basesDeBande[bande];
  for (let offset = 0; offset <= 6; offset++) {
    ws.getCell(`${col}${base + offset}`).value = "";
  }
}

/**
 * Écrit les valeurs des étiquettes à imprimer dans la feuille modèle déjà
 * chargée (mutation en place). Les emplacements au-delà du nombre
 * d'étiquettes demandées sont vidés.
 *
 * @param {ExcelJS.Workbook} workbook classeur avec la feuille modèle déjà chargée
 * @param {Array<{etiquette: string, quantite: number}>} resultats
 * @param {{date: string, heure: string, ligne: string, equipe: string, production: string, article: string}} infos
 */
export function remplirClasseur(workbook, resultats, infos) {
  const ws = chargerFeuilleModele(workbook);
  ws.name = "Etiquettes";
  const basesDeBande = detecterBasesDeBande(ws);

  // Format papier A4 explicite : l'extraction de la feuille modèle depuis le
  // classeur d'origine ne conserve pas les réglages d'imprimante
  // (printerSettings.bin, propres au poste qui a créé le fichier et non
  // portables). Sans ça, un lecteur/imprimante par défaut sur un autre
  // format (Letter, plus court) fait déborder le contenu sur une 2e page.
  ws.pageSetup.paperSize = PAPIER_A4;

  // Marges gauche/droite reprises telles quelles depuis Page Setup du
  // classeur d'origine (voir MARGES_HORIZONTALES). Les marges verticales
  // (top/bottom/header/footer) ne sont pas touchées : on part de celles déjà
  // présentes dans le modèle et on ne modifie que left/right.
  ws.pageSetup.margins = { ...ws.pageSetup.margins, ...MARGES_HORIZONTALES };

  // Pas de mise à l'échelle automatique : impression à 95 % de la taille
  // réelle ("Ajust to: 95% normal size").
  ws.pageSetup.fitToPage = false;
  ws.pageSetup.scale = 95;

  // Largeurs de colonnes reprises telles quelles depuis les propriétés de
  // colonnes du classeur d'origine (voir LARGEURS_COLONNES) : A, B et C
  // gardent chacune leur propre largeur, sans uniformisation.
  COLONNES.forEach((col) => {
    ws.getColumn(col).width = LARGEURS_COLONNES[col];
  });

  const etiquettesAPlat = [];
  for (const r of resultats) {
    for (let i = 0; i < r.quantite; i++) etiquettesAPlat.push(r.etiquette);
  }

  if (etiquettesAPlat.length > ETIQUETTES_PAR_PAGE) {
    console.warn(
      `⚠️  ${etiquettesAPlat.length} étiquettes demandées, mais le modèle n'en prévoit que ${ETIQUETTES_PAR_PAGE}. Les étiquettes au-delà de ${ETIQUETTES_PAR_PAGE} ne seront pas imprimées.`
    );
  }

  for (let index = 0; index < ETIQUETTES_PAR_PAGE; index++) {
    if (index < etiquettesAPlat.length) {
      remplirEtiquette(ws, basesDeBande, etiquettesAPlat[index], index, infos);
    } else {
      viderEtiquette(ws, basesDeBande, index);
    }
  }

  return workbook;
}

async function chargerModele() {
  const reponse = await fetch(modeleUrl);
  const buffer = await reponse.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

/**
 * Charge le modèle et y écrit les étiquettes à imprimer.
 *
 * @param {Array<{etiquette: string, quantite: number}>} resultats
 * @param {{date: string, heure: string, ligne: string, equipe: string, production: string, article: string}} infos
 * @returns {Promise<ExcelJS.Workbook>}
 */
export async function genererClasseur(resultats, infos) {
  const workbook = await chargerModele();
  remplirClasseur(workbook, resultats, infos);
  return workbook;
}

/**
 * Génère le fichier et déclenche le téléchargement dans le navigateur.
 */
export async function telechargerClasseur(resultats, infos) {
  const workbook = await genererClasseur(resultats, infos);
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Etiquettes_${infos.date || "sans-date"}_L${infos.ligne || "?"}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
