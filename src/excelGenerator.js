import ExcelJS from "exceljs";

// Padding utilisé pour pousser "Equipe" vers la droite sur la ligne Heure/Equipe,
// comme dans le fichier Etiquettes_prélèvements.xlsm d'origine.
const PADDING = "                        ";

// Polices et alignements repris à l'identique du classeur Etiquettes_prélèvements.xlsm
// (feuille "Etiquettes V2", colonnes A/B/C) : titre et type en Arial 8 gras centré,
// les 4 lignes d'infos en Arial 10 gras alignées à gauche.
const FONT_TITRE = { name: "Arial", size: 8, bold: true };
const FONT_STANDARD = { name: "Arial", size: 10, bold: true };
const ALIGN_CENTRE = { horizontal: "center", vertical: "middle", wrapText: true };
const ALIGN_GAUCHE = { vertical: "middle", wrapText: true };

// Hauteur exacte de chaque ligne 2 à 64, extraite du XML brut de Etiquettes_
// prélèvements.xlsm. Le classeur contient plusieurs feuilles quasi identiques, mais
// seule "Etiquettes V2" est visible (les autres, dont "Etiquettes Vierge", sont
// masquées) — c'est donc elle qui s'affiche/s'imprime par défaut à l'ouverture du
// fichier, et ses valeurs priment en cas de divergence entre feuilles (une seule
// ligne diffère réellement : la ligne 38, 8.25 sur "Etiquettes V2" vs 17.25 sur la
// feuille masquée). Ce n'est PAS un motif uniforme répété : les espacements entre
// bandes varient réellement (la bande 4, par ex., a 4 lignes vides au lieu de 2) —
// les reproduire telles quelles est indispensable pour que chaque étiquette
// s'imprime exactement à la même position physique que sur le fichier d'origine
// (étiquettes pré-découpées).
// null = pas de hauteur explicite dans l'original (la ligne utilise defaultRowHeight).
const HAUTEURS_LIGNES = [
  11.25, 11.25, 11.25, null, 11.25, 11.25, 11.25, 15.75, null, // lignes 2-10  (bande 1)
  11.25, 11.25, 11.25, null, 11.25, 11.25, 11.25, 15.75, null, // lignes 11-19 (bande 2)
  11.25, 11.25, 11.25, null, 11.25, 11.25, 11.25, 19.5, 11.25, // lignes 20-28 (bande 3)
  11.25, 11.25, 11.25, null, 11.25, 11.25, 11.25, 11.25, 17.25, 8.25, 10.5, // lignes 29-39 (bande 4, espacement de 4 lignes)
  11.25, 11.25, 11.25, null, 11.25, 11.25, 11.25, 17.25, 12.75, // lignes 40-48 (bande 5)
  11.25, 11.25, 11.25, null, 11.25, 11.25, 11.25, 12, 12.75, // lignes 49-57  (bande 6)
  11.25, 11.25, 11.25, null, 11.25, 11.25, 11.25, // lignes 58-64 (bande 7, dernière — pas d'espacement après)
];

// Ligne de début de chaque bande "PRELEVEMENTS" dans la feuille de référence.
// Espacement irrégulier entre bandes 4 et 5 (11 lignes au lieu de 9) — voir ci-dessus.
const DEBUTS_BANDE = [2, 11, 20, 29, 40, 49, 58];
const BANDES_PAR_PAGE = DEBUTS_BANDE.length;
const ETIQUETTES_PAR_PAGE = BANDES_PAR_PAGE * 3;
const DERNIERE_LIGNE = 64;

const COL_INDEX = { 0: 1, 1: 2, 2: 3 }; // A=1, B=2, C=3

function creerFeuille(workbook, nom) {
  const ws = workbook.addWorksheet(nom);
  ws.columns = [{ width: 33.88671875 }, { width: 35.6640625 }, { width: 31.6640625 }];

  // Mise en page d'impression identique à Etiquettes_prélèvements.xlsm, feuille
  // "Etiquettes V2" — la SEULE feuille non masquée du classeur, donc celle qui
  // s'affiche et s'imprime réellement par défaut à l'ouverture du fichier (les
  // autres feuilles, dont "Etiquettes Vierge", sont masquées et ne sont que des
  // gabarits internes légèrement différents). Sans ça, Excel applique ses
  // marges/échelle par défaut, ce qui crée un grand espace vide en haut de page.
  ws.properties.defaultRowHeight = 14.4;
  ws.pageSetup.margins = {
    top: 0.19685039370078741,
    bottom: 0,
    left: 0.5511811023622047,
    right: 0.2362204724409449,
    header: 0.31496062992125984,
    footer: 0.31496062992125984,
  };
  ws.pageSetup.orientation = "portrait";
  ws.pageSetup.scale = 96;

  // Zone d'impression fixée à A2:C64, comme les dimensions exactes de la feuille de
  // référence : garantit que chaque feuille imprime toujours le même format (3
  // colonnes, 64 lignes), qu'elle soit entièrement remplie ou non.
  ws.pageSetup.printArea = `A2:C${DERNIERE_LIGNE}`;

  // Pré-formate la hauteur exacte de chaque ligne (2 à 64), remplie ou non, pour que
  // la zone d'impression ait toujours le même rendu physique que l'original.
  HAUTEURS_LIGNES.forEach((hauteur, i) => {
    if (hauteur !== null) ws.getRow(2 + i).height = hauteur;
  });

  return ws;
}

function remplirEtiquette(ws, etiquette, indexDansPage, infos) {
  const { date, heure, ligne, equipe, production, article } = infos;
  const dateAffichee = date ? new Date(date + "T00:00:00").toLocaleDateString("fr-FR") : "";
  const heureTexte = heure ? heure : " ";

  const bande = Math.floor(indexDansPage / 3);
  const col = COL_INDEX[indexDansPage % 3];
  const base = DEBUTS_BANDE[bande];

  const lignes = [
    { texte: "PRELEVEMENTS ", font: FONT_TITRE, align: ALIGN_CENTRE },
    { texte: etiquette.toUpperCase(), font: FONT_TITRE, align: ALIGN_CENTRE },
    { texte: "", font: FONT_TITRE, align: ALIGN_CENTRE },
    { texte: `Date : ${dateAffichee}      Ligne : ${ligne || ""}`, font: FONT_STANDARD, align: ALIGN_GAUCHE },
    {
      texte: `Heure : ${heureTexte}${PADDING}Equipe : ${equipe || ""}`,
      font: FONT_STANDARD,
      align: ALIGN_GAUCHE,
    },
    { texte: `Production : ${production || ""}`, font: FONT_STANDARD, align: ALIGN_GAUCHE },
    { texte: `Article : ${article || ""}`, font: FONT_STANDARD, align: ALIGN_GAUCHE },
  ];

  lignes.forEach(({ texte, font, align }, offset) => {
    const cell = ws.getCell(base + offset, col);
    cell.value = texte;
    cell.font = font;
    cell.alignment = align;
  });
}

/**
 * Génère le classeur Excel dans le format des étiquettes de prélèvement :
 * blocs de 7 lignes (titre / type / [ligne vide] / date-ligne / heure-equipe /
 * production / article), répétés sur 3 colonnes (A, B, C) et 7 bandes par feuille
 * (positions et espacements exacts de Etiquettes_prélèvements.xlsm, voir
 * HAUTEURS_LIGNES/DEBUTS_BANDE). Chaque feuille est bornée à A2:C64 (21 étiquettes
 * max) ; au-delà, une nouvelle feuille est créée ("Etiquettes 2", "Etiquettes 3", ...)
 * plutôt que de dépasser la ligne 64 sur la même feuille.
 *
 * @param {Array<{etiquette: string, quantite: number}>} resultats
 * @param {{date: string, heure: string, ligne: string, equipe: string, production: string, article: string}} infos
 * @returns {ExcelJS.Workbook}
 */
export function genererClasseur(resultats, infos) {
  const workbook = new ExcelJS.Workbook();

  // Liste à plat : une entrée par étiquette individuelle à imprimer
  const etiquettesAPlat = [];
  for (const r of resultats) {
    for (let i = 0; i < r.quantite; i++) etiquettesAPlat.push(r.etiquette);
  }

  if (etiquettesAPlat.length === 0) {
    creerFeuille(workbook, "Etiquettes");
    return workbook;
  }

  etiquettesAPlat.forEach((etiquette, index) => {
    const pageIndex = Math.floor(index / ETIQUETTES_PAR_PAGE);
    const indexDansPage = index % ETIQUETTES_PAR_PAGE;

    let ws = workbook.getWorksheet(pageIndex + 1);
    if (!ws) {
      const nom = pageIndex === 0 ? "Etiquettes" : `Etiquettes ${pageIndex + 1}`;
      ws = creerFeuille(workbook, nom);
    }

    remplirEtiquette(ws, etiquette, indexDansPage, infos);
  });

  return workbook;
}

/**
 * Génère le fichier et déclenche le téléchargement dans le navigateur.
 */
export async function telechargerClasseur(resultats, infos) {
  const workbook = genererClasseur(resultats, infos);
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
