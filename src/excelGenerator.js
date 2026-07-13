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

// Hauteurs de lignes d'une bande (7 lignes de contenu + 2 lignes d'espacement),
// mesurées sur la bande "propre" (lignes 2 à 10) du fichier d'origine.
const HAUTEURS_BANDE = [11.25, 11.25, 11.25, undefined, 11.25, 11.25, 11.25, 15.75, undefined];
const LIGNES_PAR_BANDE = HAUTEURS_BANDE.length;

/**
 * Génère le classeur Excel dans le format des étiquettes de prélèvement :
 * blocs de 7 lignes (titre / type / [ligne vide] / date-ligne / heure-equipe /
 * production / article), répétés sur 3 colonnes (A, B, C), avec 2 lignes
 * d'espacement entre chaque bande — mise en page identique à
 * Etiquettes_prélèvements.xlsm.
 *
 * @param {Array<{etiquette: string, quantite: number}>} resultats
 * @param {{date: string, heure: string, ligne: string, equipe: string, production: string, article: string}} infos
 * @returns {ExcelJS.Workbook}
 */
export function genererClasseur(resultats, infos) {
  const { date, heure, ligne, equipe, production, article } = infos;

  const dateAffichee = date
    ? new Date(date + "T00:00:00").toLocaleDateString("fr-FR")
    : "";
  const heureTexte = heure ? heure : " ";

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Etiquettes");
  ws.columns = [{ width: 33.88671875 }, { width: 35.6640625 }, { width: 31.6640625 }];

  // Liste à plat : une entrée par étiquette individuelle à imprimer
  const etiquettesAPlat = [];
  for (const r of resultats) {
    for (let i = 0; i < r.quantite; i++) etiquettesAPlat.push(r.etiquette);
  }

  const colIndex = { 0: 1, 1: 2, 2: 3 }; // A=1, B=2, C=3

  etiquettesAPlat.forEach((etiquette, index) => {
    const bande = Math.floor(index / 3);
    const col = colIndex[index % 3];
    const base = 2 + bande * LIGNES_PAR_BANDE;

    for (let i = 0; i < LIGNES_PAR_BANDE; i++) {
      const hauteur = HAUTEURS_BANDE[i];
      if (hauteur !== undefined) ws.getRow(base + i).height = hauteur;
    }

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
