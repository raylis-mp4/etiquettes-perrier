import ExcelJS from "exceljs";

// =============================================================================
// COTES PHYSIQUES DU PAPIER — Lyreco 151342, équivalent exact du gabarit
// Avery L7160 : planche A4, grille 3 colonnes x 7 lignes, 21 étiquettes/feuille.
// Ce sont les seules valeurs "vérité terrain" de ce fichier ; tout le reste
// (largeurs de colonnes, hauteurs de lignes, marges) en est calculé.
// =============================================================================
const PAGE_LARGEUR_MM = 210; // A4 portrait
const PAGE_HAUTEUR_MM = 297;

const LARGEUR_ETIQUETTE_MM = 63.5;
const HAUTEUR_ETIQUETTE_MM = 38.1; // = pas vertical exact : aucun espace entre deux étiquettes
const MARGE_HAUT_MM = 15.09;
const MARGE_GAUCHE_MM = 7.2;
const PAS_HORIZONTAL_MM = 66.68; // colonne à colonne
const ECART_HORIZONTAL_MM = PAS_HORIZONTAL_MM - LARGEUR_ETIQUETTE_MM; // 3,18 mm

const BANDES_PAR_PAGE = 7;
const ETIQUETTES_PAR_LIGNE = 3;
const ETIQUETTES_PAR_PAGE = BANDES_PAR_PAGE * ETIQUETTES_PAR_LIGNE; // 21

const LIGNES_PAR_ETIQUETTE = 6; // PRELEVEMENTS / type / date-ligne / heure-equipe / production / article
const HAUTEUR_LIGNE_MM = HAUTEUR_ETIQUETTE_MM / LIGNES_PAR_ETIQUETTE; // 6,35 mm par ligne, à parts égales

// =============================================================================
// CONVERSIONS mm -> unités Excel
// =============================================================================
const MM_PAR_POUCE = 25.4;
const PT_PAR_POUCE = 72;
const PX_PAR_POUCE = 96; // référence "écran 96 dpi" utilisée par Excel pour les largeurs de colonnes

// mm -> points (hauteur de ligne). 1 mm = 72/25.4 ≈ 2,835 pt.
function mmVersPoints(mm) {
  return (mm * PT_PAR_POUCE) / MM_PAR_POUCE;
}

// mm -> pouces (marges de page, en pouces dans le format xlsx).
function mmVersPouces(mm) {
  return mm / MM_PAR_POUCE;
}

// mm -> largeur de colonne Excel (unité de caractère). La largeur de colonne
// Excel n'est pas un mm ni un pixel : c'est un nombre de caractères "0" de la
// police normale (Calibri 11), avec la relation approximative
//   pixels = largeur * MDW + marge_interne
// où MDW ("Maximum Digit Width") et la marge interne sont documentées à 7 px
// et 5 px pour Calibri 11 @ 96 dpi. En pratique, cette formule "manuel"
// produit un rendu ~1,5 % trop large : vérifié en générant un fichier avec
// ces constantes, en le convertissant réellement en PDF (LibreOffice + police
// Carlito, métriquement compatible avec Calibri) et en mesurant la position
// du texte : le pas horizontal rendu était de 67,68 mm au lieu des 66,68 mm
// visés. Constantes recalibrées empiriquement sur ce test réel (MDW ≈ 7,05 px,
// marge interne ≈ 6,01 px), ce qui ramène le pas mesuré à 66,65 mm (écart de
// 0,03 mm, négligeable).
const MDW_PX = 7.05;
const MARGE_INTERNE_PX = 6.01;
function mmVersLargeurColonne(mm) {
  const px = (mm * PX_PAR_POUCE) / MM_PAR_POUCE;
  return (px - MARGE_INTERNE_PX) / MDW_PX;
}

// =============================================================================
// DIMENSIONS DÉRIVÉES
// =============================================================================
const HAUTEUR_LIGNE_PT = mmVersPoints(HAUTEUR_LIGNE_MM); // 18 pt exactement (38,1/6 mm = 6,35 mm = 0,25 in)

// Chaque étiquette est divisée en 2 sous-colonnes de largeur égale, pour que
// "Equipe" puisse être aligné à droite SANS espaces de remplissage ajoutés à
// la main : les 5 lignes centrées/alignées à gauche fusionnent les 2
// sous-colonnes, seule la ligne Heure/Equipe les garde séparées (Heure à
// gauche dans la sous-colonne gauche, Equipe aligné à droite dans la
// sous-colonne droite).
const LARGEUR_SOUS_COLONNE_MM = LARGEUR_ETIQUETTE_MM / 2; // 31,75 mm
const LARGEUR_COL_SOUS_ETIQUETTE = mmVersLargeurColonne(LARGEUR_SOUS_COLONNE_MM);
const LARGEUR_COL_ECART = mmVersLargeurColonne(ECART_HORIZONTAL_MM);

// Largeur totale réelle de la grille (3 étiquettes + 2 écarts), utilisée pour
// calculer la marge de droite par soustraction (pas de valeur devinée).
const LARGEUR_CONTENU_MM = ETIQUETTES_PAR_LIGNE * LARGEUR_ETIQUETTE_MM + 2 * ECART_HORIZONTAL_MM; // 196,86 mm
const HAUTEUR_CONTENU_MM = BANDES_PAR_PAGE * HAUTEUR_ETIQUETTE_MM; // 266,7 mm

const MARGE_DROITE_MM = PAGE_LARGEUR_MM - MARGE_GAUCHE_MM - LARGEUR_CONTENU_MM; // 5,94 mm
const MARGE_BAS_MM = PAGE_HAUTEUR_MM - MARGE_HAUT_MM - HAUTEUR_CONTENU_MM; // 15,21 mm

// =============================================================================
// MISE EN PAGE DE LA FEUILLE
// =============================================================================
// 8 colonnes : [étiquette1: gauche, droite] [écart] [étiquette2: gauche, droite] [écart] [étiquette3: gauche, droite]
const COLS_GAUCHE = { 0: 1, 1: 4, 2: 7 }; // A, D, G
const COLS_DROITE = { 0: 2, 1: 5, 2: 8 }; // B, E, H
const DERNIERE_COLONNE = 8; // H

// Les données commencent dès la ligne 1 (pas de ligne 1 "vide" avant la zone
// d'impression) : vérifié qu'une ligne 1 non utilisée avant le début réel des
// données ajoute quand même sa hauteur par défaut (15 pt ≈ 5,3 mm) au rendu
// imprimé, même quand la zone d'impression déclarée commence à la ligne 2 —
// ce qui décalait la 7ᵉ bande hors de la page. En démarrant à la ligne 1, il
// n'y a plus de ligne non comptabilisée pour introduire ce décalage.
const PREMIERE_LIGNE = 1;
const DEBUTS_BANDE = Array.from(
  { length: BANDES_PAR_PAGE },
  (_, i) => PREMIERE_LIGNE + i * LIGNES_PAR_ETIQUETTE
);
const DERNIERE_LIGNE = DEBUTS_BANDE[BANDES_PAR_PAGE - 1] + LIGNES_PAR_ETIQUETTE - 1;

const FONT_TITRE = { name: "Arial", size: 8, bold: true };
const FONT_STANDARD = { name: "Arial", size: 10, bold: true };
const ALIGN_CENTRE = { horizontal: "center", vertical: "middle", wrapText: true };
const ALIGN_GAUCHE = { horizontal: "left", vertical: "middle", wrapText: true };
const ALIGN_DROITE = { horizontal: "right", vertical: "middle", wrapText: true };

function colLettre(index) {
  let s = "";
  let n = index;
  while (n > 0) {
    const reste = (n - 1) % 26;
    s = String.fromCharCode(65 + reste) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function creerFeuille(workbook, nom) {
  const ws = workbook.addWorksheet(nom);
  ws.columns = [
    { width: LARGEUR_COL_SOUS_ETIQUETTE }, // A - étiquette 1 gauche
    { width: LARGEUR_COL_SOUS_ETIQUETTE }, // B - étiquette 1 droite
    { width: LARGEUR_COL_ECART }, // C - écart
    { width: LARGEUR_COL_SOUS_ETIQUETTE }, // D - étiquette 2 gauche
    { width: LARGEUR_COL_SOUS_ETIQUETTE }, // E - étiquette 2 droite
    { width: LARGEUR_COL_ECART }, // F - écart
    { width: LARGEUR_COL_SOUS_ETIQUETTE }, // G - étiquette 3 gauche
    { width: LARGEUR_COL_SOUS_ETIQUETTE }, // H - étiquette 3 droite
  ];

  ws.pageSetup.paperSize = 9; // A4
  ws.pageSetup.orientation = "portrait";
  ws.pageSetup.scale = 100; // calibré sur les cotes réelles : aucune échelle à corriger
  ws.pageSetup.margins = {
    top: mmVersPouces(MARGE_HAUT_MM),
    left: mmVersPouces(MARGE_GAUCHE_MM),
    bottom: mmVersPouces(MARGE_BAS_MM),
    right: mmVersPouces(MARGE_DROITE_MM),
    header: 0,
    footer: 0,
  };
  ws.pageSetup.printArea = `A${PREMIERE_LIGNE}:${colLettre(DERNIERE_COLONNE)}${DERNIERE_LIGNE}`;

  for (let ligne = PREMIERE_LIGNE; ligne <= DERNIERE_LIGNE; ligne++) {
    ws.getRow(ligne).height = HAUTEUR_LIGNE_PT;
  }

  return ws;
}

function remplirEtiquette(ws, etiquette, indexDansPage, infos) {
  const { date, heure, ligne, equipe, production, article } = infos;
  const dateAffichee = date ? new Date(date + "T00:00:00").toLocaleDateString("fr-FR") : "";

  const bande = Math.floor(indexDansPage / ETIQUETTES_PAR_LIGNE);
  const positionColonne = indexDansPage % ETIQUETTES_PAR_LIGNE;
  const colGauche = COLS_GAUCHE[positionColonne];
  const colDroite = COLS_DROITE[positionColonne];
  const base = DEBUTS_BANDE[bande];

  function ligneFusionnee(offset, texte, font, align) {
    const r = base + offset;
    ws.mergeCells(r, colGauche, r, colDroite);
    const cell = ws.getCell(r, colGauche);
    cell.value = texte;
    cell.font = font;
    cell.alignment = align;
  }

  ligneFusionnee(0, "PRELEVEMENTS", FONT_TITRE, ALIGN_CENTRE);
  ligneFusionnee(1, etiquette.toUpperCase(), FONT_STANDARD, ALIGN_CENTRE);
  ligneFusionnee(2, `Date : ${dateAffichee}      Ligne : ${ligne || ""}`, FONT_STANDARD, ALIGN_GAUCHE);

  // Heure/Equipe : seule ligne où les 2 sous-colonnes restent séparées, pour
  // qu'"Equipe" soit vraiment aligné à droite (propriété d'alignement Excel),
  // sans espace de remplissage ajouté à la main.
  const rHeureEquipe = base + 3;
  const celluleHeure = ws.getCell(rHeureEquipe, colGauche);
  celluleHeure.value = `Heure : ${heure || ""}`;
  celluleHeure.font = FONT_STANDARD;
  celluleHeure.alignment = ALIGN_GAUCHE;
  const celluleEquipe = ws.getCell(rHeureEquipe, colDroite);
  celluleEquipe.value = `Equipe : ${equipe || ""}`;
  celluleEquipe.font = FONT_STANDARD;
  celluleEquipe.alignment = ALIGN_DROITE;

  ligneFusionnee(4, `Production : ${production || ""}`, FONT_STANDARD, ALIGN_GAUCHE);
  ligneFusionnee(5, `Article : ${article || ""}`, FONT_STANDARD, ALIGN_GAUCHE);
}

// =============================================================================
// VÉRIFICATION AUTOMATIQUE DE LA CALIBRATION
// =============================================================================
// Recalcule, à partir des hauteurs de lignes RÉELLEMENT écrites dans le
// fichier généré (pas à partir des constantes ci-dessus), que chaque bloc de
// 6 lignes fait bien 38,1 mm, et que la marge du haut correspond bien à
// 15,09 mm. Log un avertissement clair en cas d'écart (tolérance 0,05 mm,
// pour absorber les arrondis flottants) — pour détecter un problème de
// calibration dès la génération, plutôt qu'à l'impression.
const TOLERANCE_MM = 0.05;

export function verifierCalibration(workbook) {
  const problemes = [];

  workbook.worksheets.forEach((ws) => {
    const margeHautMm = ws.pageSetup.margins.top * MM_PAR_POUCE;
    if (Math.abs(margeHautMm - MARGE_HAUT_MM) > TOLERANCE_MM) {
      problemes.push(
        `[${ws.name}] marge du haut = ${margeHautMm.toFixed(3)} mm (attendu ${MARGE_HAUT_MM} mm)`
      );
    }

    const margeGaucheMm = ws.pageSetup.margins.left * MM_PAR_POUCE;
    if (Math.abs(margeGaucheMm - MARGE_GAUCHE_MM) > TOLERANCE_MM) {
      problemes.push(
        `[${ws.name}] marge de gauche = ${margeGaucheMm.toFixed(3)} mm (attendu ${MARGE_GAUCHE_MM} mm)`
      );
    }

    if (ws.pageSetup.scale !== 100) {
      problemes.push(`[${ws.name}] échelle = ${ws.pageSetup.scale}% (attendu 100%)`);
    }

    DEBUTS_BANDE.forEach((base, i) => {
      let totalPt = 0;
      for (let offset = 0; offset < LIGNES_PAR_ETIQUETTE; offset++) {
        totalPt += ws.getRow(base + offset).height;
      }
      const totalMm = (totalPt * MM_PAR_POUCE) / PT_PAR_POUCE;
      if (Math.abs(totalMm - HAUTEUR_ETIQUETTE_MM) > TOLERANCE_MM) {
        problemes.push(
          `[${ws.name}] bande ${i + 1} (lignes ${base}-${base + LIGNES_PAR_ETIQUETTE - 1}) = ` +
            `${totalMm.toFixed(3)} mm (attendu ${HAUTEUR_ETIQUETTE_MM} mm)`
        );
      }
    });
  });

  if (problemes.length > 0) {
    console.warn(
      "⚠️  Calibration des étiquettes : écart détecté par rapport aux cotes du gabarit :\n" +
        problemes.map((p) => `  - ${p}`).join("\n")
    );
  }

  return problemes;
}

/**
 * Génère le classeur Excel au format des étiquettes de prélèvement, calibré
 * sur les cotes officielles du gabarit Lyreco 151342 / Avery L7160 (3 x 7,
 * 21 étiquettes/feuille A4). Chaque feuille est bornée à 21 étiquettes ; au-delà,
 * une nouvelle feuille est créée ("Etiquettes 2", "Etiquettes 3", ...).
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
  } else {
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
  }

  verifierCalibration(workbook);

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
