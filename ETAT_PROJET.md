# Etat du projet — a reprendre ici

## PHASE 2 (nov. 2026) : personnalisation BEFA Consulting — Business Plan PDF

Objectif valide avec l'utilisateur le 14/11/2026 : produire un Business Plan **PDF** au format du cabinet BEFA
(modeles dans `Model Business Plan/`), a partir d'un formulaire + fichiers importes. Deux variantes :
- `nouveau_projet` (modele COREAL FOOD, bilan d'ouverture) ;
- `entreprise_existante` (modele dossier de credit O'IVOIRE, analyse bilan N / N-1).

Decisions actees :
- Google Trends : le personnel BEFA exporte lui-meme les CSV depuis trends.google.com (multiTimeline, geoMap,
  relatedQueries) et les depose sur la page. Pas de pytrends / SerpApi. Facultatif.
- Chiffres : toujours calcules par le moteur Python `befa/` (jamais par l'IA). L'IA redige le texte et peut
  chercher des statistiques sectorielles sur le web **avec sources fiables citees** (accord utilisateur).
- Charte : logo repris du site (`assets/branding/logo_befa.png`), couleurs du logo : orange `#F27F24`, gris `#9F999B`.
  Pied de page des modeles : "Réalisé par Bureau d'Etudes Financières et d'Assistance-Consulting — Page N".
- PDF : document Word genere puis converti par LibreOffice (deja requis) -> le consultant recoit aussi le .docx.

Lots :
1. **Moteur financier BEFA — FAIT** (`befa/prets.py`, `befa/projection.py`, `befa/historique.py`,
   `tests/test_befa_finance.py`, 16 tests). Reproduit au franc pres le compte d'exploitation 6 ans, le ratio
   de couverture et le plan de financement (annee 0/1) de COREAL, et l'analyse bilan + CR 2025/2024 d'O'IVOIRE.
   Echeanciers : annuites constantes, differe total (interets capitalises) / partiel, in fine (spot), periodicites.
   Constat : les modeles BEFA contiennent des incoherences (VA 2027-2029 d'O'IVOIRE fausse ; frais financiers
   COREAL non reproductibles depuis les conditions de pret annoncees) -> le moteur les evite par construction.
2. **Import des liasses SYSCOHADA — FAIT** (`befa/liasse.py`, `befa/ia.py`, route `POST /api/befa/liasse`
   dans `backend/befa_api.py`, `tests/test_befa_liasse.py`, 9 tests). pypdf repere les pages utiles (ident.,
   bilan actif, passif, CR) -> Claude (`claude-opus-5-5`, JSON impose, repli serveur `fallbacks: "default"`)
   recopie chaque ligne par code REF -> conversion DETERMINISTE + controles contre les totaux imprimes
   (AZ, BK, BZ, CP, DZ, XB, XC, XD, XE, XI) ; convention de signe des charges deduite de XC.
   Teste en reel sur ETTS 2025 et bilan d'ouverture COREAL 2024 (zero ecart, ~35 s par liasse) ; ces lectures
   reelles sont les fixtures `tests/fixtures/liasse_*.json`. La "note_ia" (remarque libre) n'est PAS une alerte :
   elle s'est deja trompee une fois (CP ETTS juge faux a tort). Alerte BEFA : "Autres dettes" (DM) > 20 % du
   bilan -> suggerer un reclassement en compte courant d'associes (capitaux propres), jamais d'office.
   uvicorn --reload se fige sous Windows/Python 3.14 (2 fois) : `--reload` retire de `lancer_local.bat` ;
   apres une modif de code, REDEMARRER le serveur (sinon les nouvelles routes repondent 405).
3. **Google Trends — FAIT** (`befa/trends.py`, `befa/graphiques.py`, route `POST /api/befa/trends`,
   `tests/test_befa_trends.py`, 11 tests). Lit multiTimeline / geoMap / relatedQueries / relatedEntities,
   FR ou EN, `,` ou `;`, UTF-8 / UTF-16 / cp1252, "<1". Calcule sans IA : tendance (1re vs derniere annee,
   seuil +/-15 %), pente, pic, saisonnalite (>= 2 ans), variation recente, fiabilite (> 10 % de zeros =
   faible volume). `faits()` = phrases chiffrees pour l'IA / le PDF ; `source()` et `AVERTISSEMENT` (indice
   relatif, pas un volume). Fixture REELLE : `tests/fixtures/trends_multiTimeline_huile_de_palme.csv`
   (export utilisateur : baisse -48 %, 14 % de zeros, pic 22/10/2023, mars/mai/fevrier forts).
   Graphiques matplotlib : palette categorielle validee (script dataviz) #E8711A,#1F6FB2,#3A8A4A,#8E4FA6,#B8860B.
4. **Redaction — FAIT, essai reel complet a refaire** (`befa/dossier.py`, `befa/commentaires.py`,
   `befa/redaction.py`, route `POST /api/befa/redaction`, `tests/test_befa_redaction.py`, 13 tests).
   - `DossierBEFA` = modele de toutes les saisies (2 variantes) ; exemple complet : `data/befa_exemple_coreal.json`.
   - Commentaires des tableaux financiers ecrits par le PROGRAMME (`commentaires.py`, style BEFA, "16,08 milliards
     FCFA", "en année 1" / "en 2026") : CE, financement, plan, couverture ; bilan + CR historiques (O'IVOIRE).
   - Recherche web : outil `web_search_20250305` VOLONTAIREMENT (la variante 20260209 a filtrage dynamique ne
     renvoie aucune citation dans le texte -> controle des sources impossible). Pas de `user_location` (CI refuse,
     erreur 400). Fragments cites regroupes en phrases ; affirmation sans citation = ecartee ; extraits conserves.
     Essai reel : 48 faits / 32 sources en 66 s (fixture `redaction_coreal_recherche.json`).
   - Redaction : JSON impose (sections par cle + SWOT), effort high ; renvois [Sx] valides, inconnus retires ;
     tout nombre absent des donnees fournies -> alerte consultant (`verifier_chiffres`, tolerance 0,6 %).
     Fixture reelle `redaction_coreal_brut.json` (10 sections, 12 risques, SWOT) = redaction faite SANS faits
     sources (1er essai) : zero chiffre invente detecte.
   - Correction moteur : le capital du credit spot n'entre plus dans le service de la dette (DSCR).
   - **A FAIRE** : credit API epuise le 14/11/2026 pendant l'essai reel -> relancer `rediger()` complet sur
     l'exemple COREAL (recherche + redaction avec renvois [Sx]) une fois le compte recharge.
5. **Document Word + PDF — FAIT** (`befa/document.py`, `befa/pdf.py`, routes `POST /api/befa/generer`
   (multipart : `dossier` JSON, `redaction` JSON facultative, `fichiers_trends`) et
   `GET /api/befa/telecharger/{job}/{pdf|docx}`, `tests/test_befa_document.py`, 5 tests).
   - python-docx -> LibreOffice (profil temporaire isole par conversion). Structure COREAL (8 parties) /
     O'IVOIRE (4 parties) ; page de garde + logo + contacts BEFA ; pied "Réalisé par Bureau d'Études
     Financières et d'Assistance-Consulting — Page N" ; tableaux orange, zeros en "–" ; graphiques ;
     renvois [Sx] -> appels de note [n] + page "Sources" ; sections IA absentes -> encart "À compléter".
   - Sommaire en 2-3 passes (LibreOffice ne pagine pas un sommaire Word) : 1re passe avec "§§" pour
     reperer les pages du sommaire, puis reperage des titres dans le PDF (pypdf). Teste titre par titre.
   - Pieges regles : largeurs de colonnes = grille du tableau (`_grille`, sinon LibreOffice les ignore) ;
     police du theme Word retiree des styles de titres (sinon Calibri au lieu d'Arial).
   - Exemples : `python -m scripts.generer_exemple_befa` -> `sorties/` (ignore par git) ;
     `data/befa_exemple_oivoire.json` (entreprise existante). COREAL : 20 pages en ~13 s.
   - Dockerfile : ajout de `libreoffice-writer` (conversion Word -> PDF).
6. **Formulaire web — FAIT** (`web/index.html` = page d'accueil ; ancien outil Excel deplace en
   `web/excel.html`). Routes ajoutees : `GET /api/befa/exemples/{coreal|oivoire}`, `POST /api/befa/projection`
   (apercu instantane, repond toujours 200 avec `erreurs`), montage `/assets`. `tests/test_befa_web.py` (4 tests).
   - 8 etapes : type, entreprise, projet, documents comptables (import liasse + validation chiffres + case
     "verifie"), Google Trends, donnees financieres (apercu recalcule a la saisie), redaction IA (lancement,
     relecture/correction des sections et du SWOT, sources), generation PDF + Word.
   - Etat dans localStorage (`befa_dossier_v1`, try/catch) + enregistrer/ouvrir un dossier .json (inclut la
     redaction `_redaction` et les CSV Trends `_trends_csv`) ; les cles `_...` ne sont jamais envoyees au serveur.
   - TESTE DANS CHROME (14/11/2026) : exemple COREAL, saisie -> apercu recalcule, generation 16 pages depuis le
     navigateur, relecture redaction persistee, message clair "credit API epuise", zero erreur console.
   - Non teste en reel faute de credit API : import de liasse et redaction depuis l'interface (routes testees).

## Essai reel complet — FAIT (14/11/2026, apres rechargement du credit API)
Par les routes du serveur (memes appels que le formulaire), exemple COREAL, resultats dans `sorties/essai_reel/` :
liasse ETTS 40 s (controles OK) ; Trends 0,3 s ; redaction 206 s (10 sections, 46 faits sources, 33 sources
citees, 92 renvois [Sx], ZERO chiffre non verifie) ; generation 11 s (23 pages, sommaire complet).
Renvois controles par echantillon : chiffres du texte = faits de la source citee. Correction : titres de
sources encodes en URL ("G%C3%B4h") decodes.

Wikipedia exclue de la recherche web a la demande de l'utilisateur (`blocked_domains` de `OUTIL_RECHERCHE`,
befa/redaction.py ; verifie par un appel reel : parametre accepte par la variante 20250305). Pour exclure
d'autres sites, ajouter leur domaine a cette liste.

## Reste a faire (phase 2)
- Faire valider par BEFA la mise en page du PDF (`sorties/essai_reel/Business_Plan_COREAL_essai_reel.pdf`).
- Faire valider par BEFA la mise en page du PDF (exemples dans `sorties/` via `python -m scripts.generer_exemple_befa`).
- Commit des phases 1 (modifs non commitees anterieures) et 2 : a faire sur demande de l'utilisateur.

---

## PHASE 1 (archive) — classeur Excel

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
