# Etat du projet — a reprendre ici

Derniere session : Lot 1, Lot 2, Lot 3 termines et testes (via API, verification navigateur reel non faite faute d'extension Chrome connectee dans cette session).

## Fait et valide
- **Lot 1** : gabarit restructure (`model/Business_Plan_Template.xlsx`), moteur `engine/`, tests verts.
- **Lot 2** : backend FastAPI, frontend fidele a la maquette, teste de bout en bout dans un vrai navigateur.
- **Lot 3** : apercu par feuille avant telechargement.
  - `engine/preview.py` : conversion feuille -> HTML, teste par `tests/test_preview.py` (11 feuilles, aucune formule brute, valeurs recalculees presentes, format FCFA preserve).
  - `backend/app.py` : `POST /api/preview` (valide + genere + recalcule + convertit en HTML, garde le fichier recalcule en cache serveur par `job_id` dans `_JOBS`) et `GET /api/download/{job_id}` (sert le fichier deja recalcule, aucun second recalcul). `POST /api/generate` (Lot 2) conserve pour compatibilite.
  - `web/index.html` : le bouton "Générer" appelle desormais `/api/preview`, affiche des onglets pour les 11 feuilles (`.preview-tab` / `.preview-sheet`), et le bouton de telechargement final (`downloadFinal()` -> `/api/download/{job_id}`) n'apparait qu'apres un apercu reussi.
  - 24 tests verts (`pytest tests/`), y compris chaine complete `/api/preview` -> `/api/download/{job_id}` verifiee via requetes HTTP directes (11 feuilles retournees, telechargement .xlsx valide, 86 Ko).

## A faire au prochain lot
- **Verification visuelle navigateur reel** : non effectuee cette session (extension Chrome non connectee). A refaire au demarrage de la prochaine session : lancer `uvicorn backend.app:app --reload --port 8000`, ouvrir `http://127.0.0.1:8000/`, remplir le formulaire, verifier que les onglets s'affichent correctement (mise en forme, pas de debordement horizontal) et que le telechargement final fonctionne.
- **Lot 4** : integration IA pour les textes narratifs (`description_activite`, `justification_croissance`) et, si retenu, suggestions de valeurs par defaut sectorielles (cf. cahier des charges §2 et §4).

## Pour reprendre
Dire "continue" ou "verifie le Lot 3 dans un navigateur puis passe au Lot 4" suffit — ce fichier et le code deja commite donnent tout le contexte necessaire.
