# F10 Facturation

Site de facturation mobile (Vite/TypeScript), formulaire d'interventions, total hebdomadaire TPS/TVQ, export Excel 1 ou 4 semaines, partage Android avec téléchargement de secours.

## Installation

```bash
npm install
npm run dev
```

**Gabarit obligatoire** : placer le fichier Excel original `modele_original.xlsx` dans `public/resources/` avant de lancer le site. Il n'est pas inclus dans ce dépôt pour le moment. Le site utilisera ce fichier comme modèle pour chaque exportation.

**Données** : enregistrées dans le stockage local du navigateur uniquement. Ne pas utiliser pour des informations sensibles ni pour la facturation en production sans sauvegarde, authentification et validation des exports.

**Partage** : si Android refuse le partage natif, le fichier Excel est téléchargé pour être joint manuellement à Gmail.

**Attention** : les fichiers XLSX générés n'ont pas encore été validés par le client.

Redéploiement Cloudflare
