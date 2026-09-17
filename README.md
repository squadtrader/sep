# Projections des banques centrales — collecte automatisée

Fait tourner les 4 scrapers (Fed, BOC, BCE, BoE) directement sur les serveurs
de GitHub Actions, plutôt que sur ta machine — ça évite tout problème de
réseau, proxy ou antivirus local.

## Mise en place (une seule fois)

1. Crée un nouveau dépôt GitHub (public ou privé), ou réutilise un dépôt
   existant.
2. Copie dedans la structure de ce dossier telle quelle :
   ```
   scripts/fed_sep_scraper.py
   scripts/boc_mpr_scraper.py
   scripts/ecb_projections_scraper.py
   scripts/boe_mpr_scraper.py
   .github/workflows/central-bank-projections.yml
   requirements.txt
   ```
3. Committe et pousse (`git add . && git commit -m "Setup" && git push`).
4. Dans les paramètres du dépôt : **Settings → Actions → General →
   Workflow permissions**, sélectionne **"Read and write permissions"**
   (nécessaire pour que le workflow puisse committer les résultats).

## Utilisation

- **Manuellement** : onglet **Actions** du dépôt → "Projections des banques
  centrales" → bouton **Run workflow**.
- **Automatiquement** : le workflow tourne aussi tous les 1er du mois. Comme
  chaque script ne collecte que les rapports déjà publiés, ça ne pose aucun
  souci de le faire tourner plus souvent que les publications elles-mêmes.

## Résultats

Après chaque exécution, les fichiers sont committés directement dans le
dépôt sous :
```
data/fed/fed_sep.json      + fed_sep.xlsx
data/boc/boc_mpr.json      + boc_mpr.xlsx
data/ecb/ecb_projections.json + ecb_projections.xlsx
data/boe/boe_mpr.json      + boe_mpr.xlsx
```
Tu peux les télécharger directement depuis GitHub (bouton "Download" sur la
page du fichier), ou cloner/puller le dépôt pour les récupérer localement.

## Si un des scrapers échoue

Chaque étape a `continue-on-error: true` : si un site change de structure
(ça arrive, comme on l'a vu avec la BOC et la BCE), les 3 autres scripts
tournent quand même et leurs résultats sont committés. Le détail de l'erreur
est visible dans les logs de l'étape correspondante, onglet Actions du
dépôt — copie-le-moi si besoin de corriger le script concerné.
