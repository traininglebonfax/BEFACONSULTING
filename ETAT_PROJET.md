# Etat du projet — a reprendre ici

Derniere session : Lot 1, Lot 2 termines et testes. Lot 3 (apercu par feuille) EN COURS.

## Fait et valide
- **Lot 1** : gabarit restructure (`model/Business_Plan_Template.xlsx`), moteur `engine/`, 16 tests verts (`tests/test_engine.py`), README a jour. Un bug de format de nombre (cellules % affichees en FCFA) trouve et corrige en cours de Lot 3 — le gabarit a ete regenere, tests revalidés apres correction.
- **Lot 2** : backend FastAPI (`backend/app.py`, endpoint `/api/generate`), frontend fonctionnel (`web/index.html`) fidele a la maquette, teste de bout en bout dans un vrai navigateur (formulaire -> generation -> telechargement).

## Lot 3 — ce qui reste a faire
1. `engine/preview.py` est cree (conversion feuille -> HTML, lecture des valeurs recalculees, couleurs/gras/formats FCFA preserves) mais **pas encore teste**.
2. Backend : ajouter un endpoint `POST /api/preview` (valide + genere + recalcule + convertit en HTML, garde le fichier en cache serveur par `job_id`) et `GET /api/download/{job_id}` (sert le fichier deja recalcule, pas de double recalcul). Le `/api/generate` actuel du Lot 2 peut rester pour compatibilite.
3. Frontend : apres soumission, appeler `/api/preview` au lieu de telecharger directement ; afficher des onglets pour naviguer entre les 11 feuilles (contenu HTML retourne) ; le bouton de telechargement final n'apparait qu'apres apercu reussi (zero erreur de formule) — cf. cahier des charges §5.
4. Tester `engine/preview.py` unitairement (au moins verifier que le HTML produit contient les bonnes valeurs et pas de formule brute).
5. Verifier visuellement dans un navigateur reel (comme pour le Lot 2).

## Pour reprendre
Dire "continue le Lot 3" suffit — ce fichier et le code déjà commité donnent tout le contexte necessaire.
