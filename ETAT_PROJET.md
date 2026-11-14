# Etat du projet — a reprendre ici

Derniere session de la phase 1 : Lots 1, 2, 3 et 4 termines et testes (Lot 4 teste avec un vrai appel API Anthropic).

## Fait et valide
- **Lot 1** : gabarit restructure (`model/Business_Plan_Template.xlsx`), moteur `engine/`, tests verts.
- **Lot 2** : backend FastAPI, frontend fidele a la maquette, teste de bout en bout dans un vrai navigateur.
- **Lot 3** : apercu par feuille avant telechargement (`POST /api/preview`, `GET /api/download/{job_id}`, onglets frontend). Verification visuelle navigateur reel toujours a refaire (extension Chrome non connectee dans cette session, cf. section suivante) — la chaine complete a ete verifiee via requetes HTTP directes.
- **Lot 4** : generation IA des textes narratifs.
  - `engine/ai_text.py` : appelle l'API Anthropic (`claude-opus-5` par defaut, `ANTHROPIC_MODEL` pour surcharger) pour generer `description_activite` et `justification_croissance` a partir des donnees du formulaire (secteur, produits, profil de croissance). Leve `AiTextError` si la cle API est absente ou l'appel echoue — jamais bloquant.
  - `engine/generator.py::attach_ai_comments` : attache ces textes en **commentaires de cellule** openpyxl sur `1_HYPOTHESES!A1` (description) et `1_HYPOTHESES!A30` (justification) — decision actee avec l'utilisateur : zero modification du gabarit, aucune formule ni valeur touchee. Appele APRES le recalcul LibreOffice (les commentaires ne survivraient pas forcement a un recalcul anterieur).
  - `backend/app.py::_generate_and_recalc` : orchestre generation -> recalcul -> appel IA (best-effort) -> attache commentaires. Utilise par `/api/preview` (retourne aussi `ai_texts`/`ai_warning` dans la reponse JSON) et par `/api/generate` (Lot 2, conserve pour compatibilite).
  - `web/index.html` : affiche un encart "Genere par l'IA" avec les deux textes au-dessus des onglets d'apercu, ou un avertissement ambre si l'IA est indisponible (le Business Plan reste genere dans tous les cas).
  - `.env` (ignore par git) porte `ANTHROPIC_API_KEY`, charge via `python-dotenv` au demarrage du backend (`backend/app.py`).
  - 5 nouveaux tests (`tests/test_ai_text.py`) : erreur claire sans cle API, parsing de la reponse IA, attache des commentaires sans alterer les valeurs/formules. 25 tests verts au total (`pytest tests/`). Teste aussi en conditions reelles avec une cle API valide (appel HTTP direct + verification du fichier telecharge).
  - Suggestions de valeurs par defaut sectorielles (role optionnel du cahier des charges Section2) : **implementees**, limitees au cout variable unitaire par produit (l'exemple donne dans le cahier des charges : "taux de marge courant pour un secteur donne").
    - Champ "Cout variable unitaire" du formulaire desormais optionnel : laisse vide, il n'est pas envoye au backend (`cout_variable_unitaire: None`).
    - `engine/ai_text.py::suggest_variable_costs` : appelle Claude avec secteur + prix de vente des produits concernes, retourne {index: valeur_suggeree}. Appele AVANT `projet.validate()` dans `backend/app.py::_generate_and_recalc` ; si l'IA echoue, erreur 422 claire invitant a saisir la valeur manuellement (aucune valeur financiere n'est jamais inventee silencieusement).
    - `engine/generator.py::attach_cost_suggestion_comments` : commentaire de cellule sur `1_HYPOTHESES!B{35+i}` (meme mecanisme que les textes narratifs).
    - `engine/preview.py` : nouveau parametre `ai_suggested_cells` sur `workbook_to_preview`/`sheet_to_html` - badge violet "✦ IA" distinct du badge vert existant (texte narratif), cf. cahier des charges "badge distinct de pré-rempli fixe et de saisi par l'utilisateur".
    - `web/index.html` : 4e badge de legende "Suggéré par l'IA" (violet) ; panneau `ai-note` liste les couts suggeres avec invitation a resaisir le champ et regenerer si besoin (garantit "reste modifiable").
    - Teste en conditions reelles (cle API valide) : suggestion generee, ecrite dans le fichier telecharge, commentaire present.

## Rappel process (a ne jamais oublier)
**L'utilisateur genere une cle Anthropic valable 24h.** A chaque fois qu'il demande de lancer l'application pour un test, lui rappeler de mettre a jour `ANTHROPIC_API_KEY` dans `.env` avant de demarrer le serveur.

## A faire au prochain lot
- **Verification visuelle navigateur reel** du Lot 3 (onglets, mise en forme) : toujours pas faite, extension Chrome non connectee lors des dernieres sessions.
- Envisager de couvrir `_generate_and_recalc` avec un test end-to-end backend (actuellement teste manuellement via HTTP, pas dans `tests/`).
- Les 4 lots du cahier des charges sont maintenant tous implementes. Prochaine etape a definir avec l'utilisateur (durcissement, deploiement, ou nouvelles fonctionnalites hors perimetre initial).

## Pour reprendre
Dire "continue" suffit — ce fichier et le code deja commite donnent tout le contexte necessaire. Ne pas oublier le rappel sur la cle API ci-dessus des qu'un lancement de serveur est demande.
