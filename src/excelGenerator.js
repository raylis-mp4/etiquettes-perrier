import ExcelJS from "exceljs";

// ---------------------------------------------------------------------------
// Calibration papier : Lyreco 151342 = équivalent exact du gabarit Avery L7160
// (planche A4, 3 colonnes x 7 lignes, 21 étiquettes/feuille). Cotes officielles
// du fabricant (et non plus des valeurs extraites — possiblement mal calées —
// du fichier Etiquettes_prélèvements.xlsm d'origine) :
//   - largeur étiquette ....... 63,5 mm
//   - hauteur étiquette ....... 38,1 mm  (= pas vertical : aucun espace entre
//                                          deux étiquettes dans une même colonne)
//   - marge du haut ........... 15,09 mm
//   - marge de gauche ......... 7,2 mm
//   - pas horizontal .......... 66,68 mm (colonne à colonne)
// L'écart horizontal entre deux étiquettes (66,68 - 63,5 = 3,18 mm) est
// matérialisé par une colonne "gouttière" étroite entre chaque colonne
// d'étiquette, pour que le texte centré (titre) reste centré sur l'étiquette
// physique réelle et non sur le pas complet.
// ---------------------------------------------------------------------------

const MM_PAR_POUCE = 25.4;
const PT_PAR_POUCE = 72;
const PX_PAR_POUCE = 96; // référence Excel pour la conversion largeur de colonne <-> pixels

const LARGEUR_ETIQUETTE_MM = 63.5;
const HAUTEUR_ETIQUETTE_MM = 38.1;
const MARGE_HAUT_MM = 15.09;
const MARGE_GAUCHE_MM = 7.2;
const PAS_HORIZONTAL_MM = 66.68;
const ECART_HORIZONTAL_MM = PAS_HORIZONTAL_MM - LARGEUR_ETIQUETTE_MM; // 3,18 mm

const LIGNES_PAR_ETIQUETTE = 7; // titre / type / vide / date-ligne / heure-equipe / production / article

function mmVersPoints(mm) {
  return (mm / MM_PAR_POUCE) * PT_PAR_POUCE;
}

function mmVersPouces(mm) {
  return mm / MM_PAR_POUCE;
}

// Formule standard Excel (Calibri 11 @ 96 dpi, largeur du "0" = 7px, marge
// interne = 5px) pour convertir une largeur physique en unité de colonne Excel :
// pixels = largeur_colonne * 7 + 5  =>  largeur_colonne = (pixels - 5) / 7
function mmVersLargeurColonne(mm) {
  const px = (mm / MM_PAR_POUCE) * PX_PAR_POUCE;
  return (px - 5) / 7;
}

const LARGEUR_COL_ETIQUETTE = mmVersLargeurColonne(LARGEUR_ETIQUETTE_MM); // ≈ 33.5714
const LARGEUR_COL_ECART = mmVersLargeurColonne(ECART_HORIZONTAL_MM); // ≈ 1.0027

// Hauteur de chacune des 7 lignes d'une étiquette : le bloc doit occuper
// EXACTEMENT 38,1 mm, donc on divise à parts égales (aucune ligne vide
// superflue entre deux étiquettes consécutives dans une même colonne).
const HAUTEUR_LIGNE_PT = mmVersPoints(HAUTEUR_ETIQUETTE_MM) / LIGNES_PAR_ETIQUETTE; // = 108/7 ≈ 15.4286 pt

// Padding utilisé pour pousser "Equipe" vers la droite sur la ligne Heure/Equipe.
const PADDING = "                        ";

const FONT_TITRE = { name: "Arial", size: 8, bold: true };
const FONT_STANDARD = { name: "Arial", size: 10, bold: true };
const ALIGN_CENTRE = { horizontal: "center", vertical: "middle", wrapText: true };
const ALIGN_GAUCHE = { vertical: "middle", wrapText: true };

// Colonnes de la feuille : A=étiquette, B=gouttière, C=étiquette, D=gouttière,
// E=étiquette. Seules les colonnes A/C/E portent du texte.
const COL_INDEX = { 0: 1, 1: 3, 2: 5 };

const PREMIERE_LIGNE = 2;
const BANDES_PAR_PAGE = 7;
const ETIQUETTES_PAR_PAGE = BANDES_PAR_PAGE * 3; // 21

// Ligne de début de chaque bande : 7 lignes par étiquette, sans aucun
// espacement entre bandes (le pas vertical réel est exactement 38,1 mm).
const DEBUTS_BANDE = Array.from(
  { length: BANDES_PAR_PAGE },
  (_, i) => PREMIERE_LIGNE + i * LIGNES_PAR_ETIQUETTE
);
const DERNIERE_LIGNE = DEBUTS_BANDE[BANDES_PAR_PAGE - 1] + LIGNES_PAR_ETIQUETTE - 1; // 50

function creerFeuille(workbook, nom) {
  const ws = workbook.addWorksheet(nom);
  ws.columns = [
    { width: LARGEUR_COL_ETIQUETTE }, // A
    { width: LARGEUR_COL_ECART }, // B (gouttière)
    { width: LARGEUR_COL_ETIQUETTE }, // C
    { width: LARGEUR_COL_ECART }, // D (gouttière)
    { width: LARGEUR_COL_ETIQUETTE }, // E
  ];

  ws.pageSetup.paperSize = 9; // A4
  ws.pageSetup.orientation = "portrait";
  ws.pageSetup.scale = 100; // calibré sur les cotes réelles : pas d'échelle à corriger
  // Seules les marges haut/gauche sont dictées par le gabarit (elles fixent
  // la position de la 1ère étiquette). Les marges bas/droite n'ont aucune
  // valeur de référence à respecter : elles ne font que réserver l'espace
  // vide après la dernière étiquette. On les fixe volontairement petites
  // MAIS avec une marge de sécurité confortable par rapport à l'espace
  // réellement disponible (5,94 mm à droite, 15,21 mm en bas), pour ne
  // jamais dépendre d'un calcul au mm près qu'un arrondi Excel ou une zone
  // non imprimable d'imprimante pourrait faire déborder sur une page
  // supplémentaire.
  ws.pageSetup.margins = {
    top: mmVersPouces(MARGE_HAUT_MM),
    left: mmVersPouces(MARGE_GAUCHE_MM),
    bottom: mmVersPouces(5),
    right: mmVersPouces(2),
    header: 0,
    footer: 0,
  };

  // Zone d'impression fixée à A2:E{DERNIERE_LIGNE} : 3 colonnes d'étiquettes
  // (A/C/E) + 2 colonnes de gouttière (B/D), 7 bandes de 7 lignes.
  ws.pageSetup.printArea = `A${PREMIERE_LIGNE}:E${DERNIERE_LIGNE}`;

  for (let ligne = PREMIERE_LIGNE; ligne <= DERNIERE_LIGNE; ligne++) {
    ws.getRow(ligne).height = HAUTEUR_LIGNE_PT;
  }

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
 * Génère le classeur Excel au format des étiquettes de prélèvement, calibré
 * sur les cotes officielles du gabarit Lyreco 151342 / Avery L7160 (3 x 7,
 * 21 étiquettes/feuille A4). Chaque feuille est bornée à A2:E{DERNIERE_LIGNE}
 * (21 étiquettes max) ; au-delà, une nouvelle feuille est créée
 * ("Etiquettes 2", "Etiquettes 3", ...).
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
