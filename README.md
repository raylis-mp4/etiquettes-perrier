# Étiquettes de prélèvement — Perrier

Générateur d'étiquettes de prélèvement PET L30-L34, basé sur QUAL/IT-606/VGZ rev15 (50cL).

## Démarrage

```bash
npm install
npm run dev
```

Ouvre http://localhost:5173

## Build pour déploiement

```bash
npm run build
```
Le résultat est dans `dist/`, à déployer sur Vercel/Netlify (gratuit) ou n'importe quel hébergeur statique.

## Structure

- `src/rulesEngine.js` — le tableau de règles (quantités par situation), validé avec Youssouf
- `src/excelGenerator.js` — génère le fichier .xlsx avec ExcelJS (styles Arial 8/10, gras, centré)
- `src/App.jsx` — l'interface (formulaire + aperçu + téléchargement)

## Prochaines étapes possibles

- Ajouter Supabase pour stocker le tableau de règles en base plutôt qu'en dur dans le code
  (utile si les règles qualité changent)
- Déployer sur Vercel pour un accès direct depuis le téléphone au poste
