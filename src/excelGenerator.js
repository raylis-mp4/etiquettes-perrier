import ExcelJS from "exceljs";
import modeleUrl from "./assets/modele-etiquettes.xlsx?url";

// =============================================================================
// Générateur d'étiquettes de prélèvement — approche "modèle rempli".
//
// Le fichier src/assets/modele-etiquettes.xlsx est un export exact et non
// modifié de la feuille "Etiquettes Vierge V2" du classeur d'origine
// Etiquettes_prélèvements.xlsm (macros retirées, elles ne servaient qu'à
// faire ce que ce fichier fait déjà en JS). Hauteurs de lignes, polices et
// marges viennent de ce fichier et ne sont jamais recalculées : ce module se
// contente d'écrire des valeurs dans des cellules déjà stylées. Seules
// exceptions, corrigées après coup car le modèle d'origine n'était pas fiable
// sur ces deux points précis : le format papier (voir PAPIER_A4) et la
// largeur des colonnes (voir LARGEUR_COLONNE_ETIQUETTE).
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

// Les 3 colonnes du modèle d'origine (A=33.89, B=34.0, C=31.66) n'ont jamais
// été calibrées de façon uniforme horizontalement, contrairement aux
// hauteurs de ligne verticalement justes. Sur le papier pré-découpé Lyreco
// 151342 (équivalent Avery L7160), le pas horizontal entre les 3 colonnes
// d'étiquettes physiques est parfaitement uniforme : 66,68 mm. On force donc
// les 3 colonnes à une largeur identique calculée à partir de ce pas,
// conversion mm -> largeur de colonne Excel (largeur_px = mm * 3.78 ;
// largeur_excel = (largeur_px - 5) / 7).
const PAS_HORIZONTAL_MM = 66.68;
const LARGEUR_COLONNE_ETIQUETTE = (PAS_HORIZONTAL_MM * 3.78 - 5) / 7;

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
const TYPES_ETIQUETTE = {
  "Hydrothèque": { offset: 1, texte: " HYDROTHEQUE " },
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
  ws.getCell(`${col}${base + style.offset}`).value = style.texte;
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
  // ExcelJS réinjecte toujours fitToWidth/fitToHeight="1" à l'écriture, même
  // quand le fichier d'origine ne les avait pas : ça peut pousser certains
  // lecteurs à re-proportionner tout le rendu ("Ajuster à la page"), ce qui
  // fausse les tailles de colonnes/lignes réellement imprimées. On les
  // neutralise pour revenir strictement à l'échelle fixe (96 %) du modèle.
  ws.pageSetup.fitToWidth = undefined;
  ws.pageSetup.fitToHeight = undefined;

  // Largeurs de colonnes uniformes (voir LARGEUR_COLONNE_ETIQUETTE) : le
  // modèle d'origine a 3 largeurs différentes (A/B/C), ce qui décale de plus
  // en plus le contenu de sa case physique de la 1re à la 3e colonne sur le
  // papier pré-découpé.
  COLONNES.forEach((col) => {
    ws.getColumn(col).width = LARGEUR_COLONNE_ETIQUETTE;
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
